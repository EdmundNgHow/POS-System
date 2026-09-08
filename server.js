const crypto = require("crypto");
const path = require("path");
const fs = require("fs");
const os = require("os");
const express = require("express");
const QRCode = require("qrcode");
const XLSX = require("xlsx");
const { loadEnvFile } = require("./server/bootstrap/env");
const { createNetworkService } = require("./server/bootstrap/network");
const { createConfiguredStore } = require("./server/bootstrap/storage");
const { registerModuleApi } = require("./server/modules/api");
const { getModuleCatalog, registerPageRoutes } = require("./server/modules/registry");
const {
  allPermissionIds,
  hasPermission,
  configureRolePermissions,
  normalizeRole,
  permissionsForSettings,
  requirePermission,
  rolePermissionSnapshot,
  roleCatalog
} = require("./server/security/permissions");

loadEnvFile(path.join(__dirname, ".env"));

const app = express();
const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || "0.0.0.0";
const TAX_RATE = 0;
const STAFF_RECEIPT_STATUSES = new Set(["paid", "unpaid", "partially_paid"]);
const DEFAULT_COMPANY_LOGO = "/assets/tunas-logo.jpg";
const { store, dataFile, backupDir } = createConfiguredStore(__dirname);
const { getNetworkAccessInfo } = createNetworkService({ host: HOST, port: PORT });
if (typeof store.getSetting === "function") {
  try {
    const configuredPermissions = JSON.parse(store.getSetting("role_permissions_json", "{}") || "{}");
    configureRolePermissions(configuredPermissions);
  } catch (_) {
    configureRolePermissions({});
  }
}

const sessions = new Map();
const eventClients = new Set();

app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

function storageInfo() {
  return typeof store.getStorageInfo === "function"
    ? store.getStorageInfo()
    : { type: store.storageType || "unknown" };
}

function storageHealth() {
  try {
    if (typeof store.checkConnection === "function") {
      store.checkConnection();
    }
    return { ok: true, ...storageInfo() };
  } catch (err) {
    return { ok: false, error: err.message || "Database unavailable.", ...storageInfo() };
  }
}

function reloadBusinessData(req, res, next) {
  if (req.path === "/health" || req.path === "/network/access" || req.path === "/events") {
    return next();
  }
  try {
    if (typeof store.reloadFromDatabase === "function") {
      store.reloadFromDatabase();
    }
    return next();
  } catch (err) {
    return res.status(503).json({
      error: "Database is temporarily unavailable. Please keep this transaction open and retry when reconnected.",
      database: storageHealth()
    });
  }
}

function broadcastDataChange(type = "data_changed", payload = {}) {
  const message = JSON.stringify({
    type,
    payload,
    storage: storageInfo(),
    created_at: new Date().toISOString()
  });
  for (const client of eventClients) {
    try {
      client.write(`event: ${type}\n`);
      client.write(`data: ${message}\n\n`);
    } catch (_) {
      eventClients.delete(client);
    }
  }
}

const originalStoreSave = typeof store._save === "function" ? store._save.bind(store) : null;
if (originalStoreSave) {
  store._save = (...args) => {
    const result = originalStoreSave(...args);
    broadcastDataChange("data_changed");
    return result;
  };
}

app.use("/api", reloadBusinessData);

registerPageRoutes(app, { appRoot: __dirname });

function toMoney(cents) {
  return Number((Number(cents || 0) / 100).toFixed(2));
}

function statusFromCode(code) {
  if (code === "BAD_REQUEST") return 400;
  if (code === "FORBIDDEN") return 403;
  if (code === "NOT_FOUND") return 404;
  if (code === "CONFLICT") return 409;
  return 500;
}

function parseLogDetails(details) {
  try {
    return JSON.parse(details || "{}");
  } catch (_) {
    return { raw: String(details || "") };
  }
}

function canViewCost(user) {
  return hasPermission(user, "inventory.view_cost");
}

function isAdmin(user) {
  const role = normalizeRole(user?.role);
  return role === "admin" || role === "super_admin" || role === "system_owner";
}

function normalizeReceiptStatus(value) {
  const status = String(value || "").trim().toLowerCase().replace(/\s+/g, "_");
  if (status === "draft") return "unpaid";
  if (status === "partial" || status === "partially-paid") return "partially_paid";
  if (status === "voided") return "void";
  return status;
}

function paymentMethodLabel(value) {
  const method = String(value || "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  const labels = {
    cash: "Cash",
    duitnow_qr: "DuitNow QR",
    credit_card: "Credit Card",
    debit_card: "Debit Card",
    bank_transfer: "Bank Transfer",
    e_wallet: "E-Wallet",
    havent_paid: "Haven't Paid"
  };
  return labels[method] || String(value || "-");
}

function staffCanUpdateReceipt(payload = {}) {
  if (Object.hasOwn(payload, "items") || Object.hasOwn(payload, "user_id")) {
    return false;
  }

  if (Object.hasOwn(payload, "status")) {
    return STAFF_RECEIPT_STATUSES.has(normalizeReceiptStatus(payload.status));
  }

  return true;
}

function listFromValue(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }
  return String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
}

function mapDiscount(discount) {
  return {
    ...discount,
    is_active: Boolean(discount.is_active),
    min_purchase_amount: toMoney(discount.min_purchase_cents),
    max_discount_amount: toMoney(discount.max_discount_cents),
    promo_spend_amount: toMoney(discount.promo_spend_cents),
    promo_discount_amount: toMoney(discount.promo_discount_cents),
    applicable_categories: listFromValue(discount.applicable_categories),
    applicable_product_ids: listFromValue(discount.applicable_product_ids).map((item) => Number(item)).filter(Boolean),
    applicable_member_tiers: listFromValue(discount.applicable_member_tiers || discount.member_tier)
  };
}

function isApprovedStaffDiscount(discount) {
  if (!discount) {
    return true;
  }

  const discountId = Number(discount.id || discount.discount_id || 0);
  if (!Number.isInteger(discountId) || discountId <= 0) {
    return false;
  }

  return store.getDiscounts({ activeNow: true }).some((approved) => Number(approved.id) === discountId);
}

function mapProduct(product, user = null) {
  const unitPrice = toMoney(product.unit_price_cents);
  const cost = toMoney(product.cost_cents);
  const mapped = {
    ...product,
    brand: product.brand || "",
    supplier: product.supplier || "",
    description: product.description || "",
    template_type: product.template_type || "custom",
    attributes: parseLogDetails(product.attributes_json || "{}"),
    status: product.is_active ? "active" : "inactive",
    unit_price: unitPrice,
    selling_price: unitPrice
  };
  if (canViewCost(user)) {
    mapped.cost = cost;
    mapped.cost_price = cost;
    mapped.margin = unitPrice > 0 ? Number((((unitPrice - cost) / unitPrice) * 100).toFixed(2)) : 0;
  } else {
    delete mapped.cost_cents;
  }
  return mapped;
}

function mapUser(user) {
  return {
    id: user.id,
    username: user.username,
    full_name: user.full_name,
    role: user.role,
    email: user.email || "",
    phone: user.phone || "",
    employee_id: user.employee_id || "",
    profile_photo: user.profile_photo || "",
    status: user.status || (user.is_active ? "active" : "inactive"),
    is_active: user.status ? user.status === "active" : Boolean(user.is_active),
    force_password_change: Boolean(user.force_password_change),
    last_login_at: user.last_login_at || "",
    created_at: user.created_at || "",
    updated_at: user.updated_at || ""
  };
}

function mapSupplier(supplier) {
  return {
    ...supplier,
    outstanding_amount: toMoney(supplier.outstanding_cents),
    is_active: Boolean(supplier.is_active)
  };
}

function mapPurchaseOrder(order) {
  return {
    ...order,
    subtotal: toMoney(order.subtotal_cents),
    discount: toMoney(order.discount_cents),
    total: toMoney(order.total_cents),
    supplier: order.supplier ? mapSupplier(order.supplier) : null,
    items: (order.items || []).map((item) => ({
      ...item,
      unit_cost: toMoney(item.unit_cost_cents),
      line_total: toMoney(item.line_total_cents),
      remaining_quantity: Math.max(0, Number(item.quantity_ordered || 0) - Number(item.quantity_received || 0))
    }))
  };
}

function mapExpense(expense) {
  return {
    ...expense,
    amount: toMoney(expense.amount_cents)
  };
}

function mapDailyClosing(closing) {
  return {
    ...closing,
    opening_cash: toMoney(closing.opening_cash_cents),
    cash_sales: toMoney(closing.cash_sales_cents),
    cash_expenses: toMoney(closing.cash_expenses_cents),
    expected_cash: toMoney(closing.expected_cash_cents),
    counted_cash: toMoney(closing.counted_cash_cents),
    variance: toMoney(closing.variance_cents)
  };
}

function mapSupplierPayment(payment) {
  return {
    ...payment,
    amount: toMoney(payment.amount_cents)
  };
}

function mapFinanceSummary(summary) {
  return {
    ...summary,
    revenue: toMoney(summary.revenue_cents),
    cash_sales: toMoney(summary.cash_sales_cents),
    expenses: toMoney(summary.expense_cents),
    income: toMoney(summary.income_cents),
    outstanding_receivable: toMoney(summary.outstanding_receivable_cents),
    outstanding_payable: toMoney(summary.outstanding_payable_cents),
    unpaid_sales: toMoney(summary.unpaid_sales_cents),
    recent_supplier_payments: (summary.recent_supplier_payments || []).map(mapSupplierPayment),
    recent_expenses: (summary.recent_expenses || []).map(mapExpense),
    recent_payments: (summary.recent_payments || []).map(mapRecentPayment)
  };
}

function safeJsonParse(value, fallback) {
  try {
    return JSON.parse(String(value || ""));
  } catch (_) {
    return fallback;
  }
}

function defaultReceiptTemplateSettings() {
  const company = getCompanyProfileSettings();
  return {
    company_logo: company.company_logo || DEFAULT_COMPANY_LOGO,
    company_name: company.company_name,
    ssm_number: company.business_registration_number,
    address: company.business_address || "",
    phone: company.phone || "",
    email: company.email || "",
    website: company.website || "",
    sst_registration_number: company.sst_registration_number || "",
    header: company.company_name,
    footer: "Computer Generated Invoice",
    receipt_notes: "Thank you",
    terms_conditions: "Goods sold are not refundable unless required by law.",
    thank_you_message: "Thank you",
    paper_size: "80mm",
    font_size: 12,
    show_qr_code: false,
    show_barcode: false,
    ask_before_print: true,
    auto_print: false,
    default_print_template: "80mm",
    duitnow_static_qr: "",
    dynamic_qr_enabled: false
  };
}

function fallbackCompanyProfileSettings() {
  return {
    id: 1,
    company_logo: DEFAULT_COMPANY_LOGO,
    company_name: "TUNAS BADMINTON SPORTS CENTRE SDN. BHD.",
    business_registration_number: "1536028-U (202301042111)",
    registration_date: "24 October 2023",
    state: "Negeri Sembilan",
    nature_of_business: [
      "Convenience Stores",
      "Retail of Football, Hockey, Cricket, Baseball, Badminton, Futsal and Paintball Equipment",
      "Wholesale of a Variety of Goods Without Any Particular Specialization (N.E.C.)"
    ],
    business_address: "Lot PT 3962 & 3963, Jln Haruan 2, Pusat Komersial Oakland, 70300 Seremban, Negeri Sembilan",
    phone: "0182063324",
    email: "",
    website: "",
    sst_registration_number: "",
    created_at: "",
    updated_at: ""
  };
}

function getCompanyProfileSettings() {
  let profile = fallbackCompanyProfileSettings();
  if (typeof store.getCompanySettings === "function") {
    profile = { ...profile, ...store.getCompanySettings() };
  }
  return { ...profile, company_logo: profile.company_logo || DEFAULT_COMPANY_LOGO };
}

function defaultPrinterSettings() {
  return {
    printer_58mm: "Thermal Printer (58mm)",
    printer_80mm: "Thermal Printer (80mm)",
    printer_a4: "Office Printer",
    default_template: "80mm",
    manual_override: true
  };
}

function getReceiptTemplateSettings() {
  const company = getCompanyProfileSettings();
  const template = {
    ...defaultReceiptTemplateSettings(),
    ...safeJsonParse(store.getSetting?.("receipt_template_json", "{}"), {})
  };
  return {
    ...template,
    company_logo: company.company_logo || template.company_logo || DEFAULT_COMPANY_LOGO,
    company_name: company.company_name,
    ssm_number: company.business_registration_number,
    address: company.business_address || "",
    phone: company.phone || "",
    email: company.email || "",
    website: company.website || "",
    sst_registration_number: company.sst_registration_number || "",
    header: company.company_name
  };
}

function getPrinterSettings() {
  return {
    ...defaultPrinterSettings(),
    ...safeJsonParse(store.getSetting?.("printer_settings_json", "{}"), {})
  };
}

function defaultInvoiceSettings() {
  return {
    prefix: "Invoice",
    date_format: "YYMM",
    number_length: 3,
    reset: "monthly"
  };
}

function normalizeInvoiceSettingsPayload(payload = {}) {
  const current = getInvoiceSettings();
  const dateFormat = String(payload.date_format ?? payload.number_format ?? current.date_format ?? "YYMM").trim().toUpperCase();
  const reset = String(payload.reset ?? payload.reset_numbering ?? current.reset ?? "monthly").trim().toLowerCase();
  const prefix = String(payload.prefix ?? payload.invoice_prefix ?? current.prefix ?? "Invoice")
    .trim()
    .replace(/\s+/g, "-")
    .replace(/[^A-Za-z0-9_-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "") || defaultInvoiceSettings().prefix;
  return {
    prefix,
    date_format: ["YYMM", "DDMM", "YYYYMM", "YYYY", "NONE"].includes(dateFormat) ? dateFormat : current.date_format || "YYMM",
    number_length: Math.min(8, Math.max(3, Number(payload.number_length ?? current.number_length ?? 3) || 3)),
    reset: ["daily", "monthly", "yearly", "never"].includes(reset) ? reset : current.reset || "monthly"
  };
}

function getInvoiceSettings() {
  return {
    ...defaultInvoiceSettings(),
    ...safeJsonParse(store.getSetting?.("invoice_settings_json", "{}"), {})
  };
}

function defaultPdfExportSettings() {
  return {
    folder: "downloads",
    custom_path: ""
  };
}

function normalizePdfExportPayload(payload = {}) {
  const current = getPdfExportSettings();
  const folder = String(payload.folder ?? current.folder ?? "downloads").trim().toLowerCase();
  return {
    folder: ["downloads", "desktop", "documents", "custom"].includes(folder) ? folder : "downloads",
    custom_path: String(payload.custom_path ?? current.custom_path ?? "").trim()
  };
}

function getPdfExportSettings() {
  return {
    ...defaultPdfExportSettings(),
    ...safeJsonParse(store.getSetting?.("pdf_export_settings_json", "{}"), {})
  };
}

function normalizeTemplatePayload(payload = {}) {
  const current = getReceiptTemplateSettings();
  const boolValue = (key) => Object.prototype.hasOwnProperty.call(payload, key) ? Boolean(payload[key]) : Boolean(current[key]);
  return {
    ...current,
    company_logo: String(payload.company_logo ?? current.company_logo ?? "").trim(),
    company_name: String(payload.company_name ?? current.company_name ?? "").trim() || current.company_name,
    ssm_number: String(payload.ssm_number ?? current.ssm_number ?? "").trim(),
    address: String(payload.address ?? current.address ?? "").trim(),
    phone: String(payload.phone ?? current.phone ?? "").trim(),
    email: String(payload.email ?? current.email ?? "").trim(),
    header: String(payload.header ?? current.header ?? "").trim(),
    footer: String(payload.footer ?? current.footer ?? "").trim(),
    receipt_notes: String(payload.receipt_notes ?? current.receipt_notes ?? "").trim(),
    terms_conditions: String(payload.terms_conditions ?? current.terms_conditions ?? "").trim(),
    thank_you_message: String(payload.thank_you_message ?? current.thank_you_message ?? "").trim(),
    paper_size: ["58mm", "80mm", "a4"].includes(String(payload.paper_size || "").toLowerCase()) ? String(payload.paper_size).toLowerCase() : current.paper_size,
    font_size: Math.min(18, Math.max(9, Number(payload.font_size ?? current.font_size ?? 12))),
    show_qr_code: boolValue("show_qr_code"),
    show_barcode: boolValue("show_barcode"),
    ask_before_print: boolValue("ask_before_print"),
    auto_print: boolValue("auto_print"),
    default_print_template: ["58mm", "80mm", "a4"].includes(String(payload.default_print_template || "").toLowerCase()) ? String(payload.default_print_template).toLowerCase() : current.default_print_template,
    duitnow_static_qr: String(payload.duitnow_static_qr ?? current.duitnow_static_qr ?? "").trim(),
    dynamic_qr_enabled: boolValue("dynamic_qr_enabled")
  };
}

function normalizePrinterPayload(payload = {}) {
  const current = getPrinterSettings();
  const boolValue = (key) => Object.prototype.hasOwnProperty.call(payload, key) ? Boolean(payload[key]) : Boolean(current[key]);
  return {
    ...current,
    printer_58mm: String(payload.printer_58mm ?? current.printer_58mm ?? "").trim() || defaultPrinterSettings().printer_58mm,
    printer_80mm: String(payload.printer_80mm ?? current.printer_80mm ?? "").trim() || defaultPrinterSettings().printer_80mm,
    printer_a4: String(payload.printer_a4 ?? current.printer_a4 ?? "").trim() || defaultPrinterSettings().printer_a4,
    default_template: ["58mm", "80mm", "a4"].includes(String(payload.default_template || "").toLowerCase()) ? String(payload.default_template).toLowerCase() : current.default_template,
    manual_override: boolValue("manual_override")
  };
}

function electronKnownPath(name) {
  try {
    const electron = require("electron");
    if (electron && typeof electron === "object" && electron.app && typeof electron.app.getPath === "function") {
      return electron.app.getPath(name);
    }
  } catch (_) {
    // The web/server runtime falls back to OS home folders.
  }
  return "";
}

function knownUserFolder(folder) {
  const key = String(folder || "downloads").toLowerCase();
  const electronPath = electronKnownPath(key === "downloads" ? "downloads" : key);
  if (electronPath) return electronPath;
  const home = os.homedir();
  if (key === "desktop") return path.join(home, "Desktop");
  if (key === "documents") return path.join(home, "Documents");
  return path.join(home, "Downloads");
}

function resolvePdfExportFolder(settings = getPdfExportSettings()) {
  const normalized = normalizePdfExportPayload(settings);
  const folderPath = normalized.folder === "custom"
    ? path.resolve(normalized.custom_path || knownUserFolder("downloads"))
    : knownUserFolder(normalized.folder);
  return {
    ...normalized,
    folder_path: folderPath
  };
}

function safeFileStem(value, fallback = "Invoice") {
  return String(value || fallback)
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/-+/g, "-")
    .trim()
    .replace(/^\.+|\.+$/g, "") || fallback;
}

function uniquePdfTarget(folderPath, invoiceNumber) {
  const stem = safeFileStem(invoiceNumber, "Invoice");
  let index = 0;
  while (index < 1000) {
    const suffix = index ? ` (${index})` : "";
    const fileName = `${stem}${suffix}.pdf`;
    const filePath = path.join(folderPath, fileName);
    if (!fs.existsSync(filePath)) {
      return { file_name: fileName, file_path: filePath };
    }
    index += 1;
  }
  const err = new Error("Unable to find an available PDF file name.");
  err.code = "CONFLICT";
  throw err;
}

function mapOutstandingOrder(order) {
  const payableCents = Number(order.payable_cents ?? order.total_cents);
  const paidCents = Number(order.paid_cents || 0);
  return {
    ...order,
    subtotal: toMoney(order.subtotal_cents),
    total: toMoney(payableCents),
    amount_paid: toMoney(paidCents),
    outstanding_balance: toMoney(order.outstanding_cents),
    payment_status: normalizeReceiptStatus(order.status || "unpaid")
  };
}

function mapRecentPayment(payment) {
  return {
    ...payment,
    amount: toMoney(payment.amount_cents),
    payment_method: String(payment.payment_method || ""),
    payment_status: normalizeReceiptStatus(payment.status || "")
  };
}

function mapOrderForApi(order) {
  if (!order) {
    return null;
  }
  const payableCents = Number(order.payable_cents ?? order.total_cents);
  const paidCents = Number(order.paid_cents || 0);
  return {
    ...order,
    invoice_number: String(order.invoice_number || ""),
    subtotal: toMoney(order.subtotal_cents),
    tax: toMoney(order.tax_cents),
    discount: toMoney(order.discount_cents),
    total: toMoney(payableCents),
    raw_total: toMoney(order.total_cents),
    rounding: toMoney(order.rounding_cents),
    paid: toMoney(paidCents),
    change: toMoney(paidCents - payableCents),
    outstanding_balance: toMoney(Math.max(0, payableCents - paidCents)),
    payment_status: normalizeReceiptStatus(order.status || "paid"),
    payment_method: order.payment_method || "",
    user: order.user || null,
    member: order.member || null,
    items: (order.items || []).map((item) => ({
      ...item,
      unit_price: toMoney(item.unit_price_cents),
      line_total: toMoney(item.line_total_cents)
    }))
  };
}

function mapCustomerOutstanding(result) {
  return {
    ...result,
    total_credit: toMoney(result.total_credit_cents),
    total_debit: toMoney(result.total_debit_cents),
    outstanding_balance: toMoney(result.outstanding_cents),
    invoices: (result.invoices || []).map((invoice) => ({
      ...invoice,
      original_amount: toMoney(invoice.original_cents),
      amount_paid: toMoney(invoice.paid_cents),
      outstanding_balance: toMoney(invoice.balance_cents),
      payment_status: normalizeReceiptStatus(invoice.status || "unpaid")
    }))
  };
}

function mapCustomerOutstandingReportRow(row) {
  return {
    ...row,
    total_credit: toMoney(row.total_credit_cents),
    total_debit: toMoney(row.total_debit_cents),
    current_outstanding: toMoney(row.current_outstanding_cents)
  };
}

function mapCustomerLedgerRow(row) {
  return {
    ...row,
    debit: toMoney(row.debit_cents),
    credit: toMoney(row.credit_cents),
    running_balance: toMoney(row.running_balance_cents)
  };
}

function mapCustomerLedgerSummary(summary) {
  return {
    total_credit: toMoney(summary.total_credit_cents),
    total_debit: toMoney(summary.total_debit_cents),
    total_outstanding: toMoney(summary.total_outstanding_cents)
  };
}

function mapSoldItemsReport(report, includeFinancials = false) {
  const mapMoneyFields = (row, fields) => {
    const mapped = { ...row };
    for (const [source, target] of fields) {
      mapped[target] = toMoney(row[source]);
    }
    return mapped;
  };

  const itemRows = (report.item_rows || []).map((row) => {
    const mapped = mapMoneyFields(row, [
      ["unit_price_cents", "unit_price"],
      ["discount_cents", "discount"],
      ["net_total_cents", "total_amount"],
      ["line_total_cents", "gross_amount"]
    ]);
    if (includeFinancials) {
      mapped.cost_price = toMoney(row.unit_cost_cents);
      mapped.cost_total = toMoney(Number(row.unit_cost_cents || 0) * Number(row.quantity || 0));
      mapped.profit = toMoney(Number(row.net_total_cents || 0) - Number(row.unit_cost_cents || 0) * Number(row.quantity || 0));
      mapped.margin_percent =
        Number(row.net_total_cents || 0) > 0
          ? Number((((Number(row.net_total_cents || 0) - Number(row.unit_cost_cents || 0) * Number(row.quantity || 0)) / Number(row.net_total_cents || 0)) * 100).toFixed(2))
          : 0;
    }
    return mapped;
  });

  const productSummary = (report.product_summary || []).map((row) => {
    const mapped = mapMoneyFields(row, [
      ["revenue_cents", "revenue"],
      ["discount_cents", "discount"],
      ["average_selling_price_cents", "average_selling_price"]
    ]);
    if (includeFinancials) {
      mapped.cost_total = toMoney(row.cost_cents);
      mapped.profit = toMoney(row.profit_cents);
      mapped.margin_percent = row.margin_percent;
    }
    return mapped;
  });

  return {
    ...report,
    totals: {
      ...report.totals,
      revenue: toMoney(report.totals?.revenue_cents),
      discount: toMoney(report.totals?.discount_cents)
    },
    item_rows: itemRows,
    category_summary: (report.category_summary || []).map((category) => ({
      ...category,
      revenue: toMoney(category.revenue_cents),
      discount: toMoney(category.discount_cents),
      products: (category.products || []).map((product) => ({
        ...product,
        revenue: toMoney(product.revenue_cents),
        discount: toMoney(product.discount_cents)
      }))
    })),
    product_summary: productSummary,
    top_products: (report.top_products || []).map((product) => ({
      ...product,
      revenue: toMoney(product.revenue_cents)
    })),
    low_stock: productSummary.filter((product) => product.low_stock),
    payment_summary: {
      cash: toMoney(report.payment_summary?.cash_cents),
      qr: toMoney(report.payment_summary?.qr_cents),
      card: toMoney(report.payment_summary?.card_cents),
      bank_transfer: toMoney(report.payment_summary?.bank_transfer_cents),
      e_wallet: toMoney(report.payment_summary?.e_wallet_cents),
      unpaid: toMoney(report.payment_summary?.unpaid_cents),
      outstanding_amount: toMoney(report.payment_summary?.outstanding_cents)
    },
    staff_summary: (report.staff_summary || []).map((staff) => ({
      ...staff,
      revenue: toMoney(staff.revenue_cents),
      discount: toMoney(staff.discount_cents)
    })),
    include_financials: includeFinancials
  };
}

function soldItemsQuery(req) {
  const query = req.query || {};
  const result = {
    preset: String(query.preset || "today").trim(),
    start_date: String(query.start_date || "").trim(),
    end_date: String(query.end_date || "").trim(),
    category: String(query.category || "").trim(),
    brand: String(query.brand || "").trim(),
    cashier: String(query.cashier || "").trim(),
    payment_method: String(query.payment_method || "").trim(),
    search: String(query.search || "").trim()
  };

  if (!isAdmin(req.auth_user)) {
    result.preset = "today";
    result.start_date = "";
    result.end_date = "";
  }

  return result;
}

function reportTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return String(value || "");
  }
  return date.toLocaleTimeString("en-MY", { hour: "2-digit", minute: "2-digit" });
}

function buildSoldItemsWorkbook(report) {
  const workbook = XLSX.utils.book_new();
  const sheetRows = (rows) => XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(
    workbook,
    sheetRows((report.item_rows || []).map((row) => ({
      Time: reportTime(row.created_at),
      "Receipt Number": row.receipt_number,
      "Product Name": row.product_name,
      Barcode: row.barcode,
      SKU: row.sku,
      Category: row.category,
      Brand: row.brand,
      "Quantity Sold": row.quantity,
      "Unit Price": row.unit_price,
      Discount: row.discount,
      "Total Amount": row.total_amount,
      Cashier: row.cashier,
      Customer: row.customer || ""
    }))),
    "Sold Items"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    sheetRows((report.product_summary || []).map((row) => ({
      "Product Name": row.product_name,
      Category: row.category,
      "Quantity Sold": row.quantity,
      Revenue: row.revenue,
      Discount: row.discount,
      "Average Selling Price": row.average_selling_price,
      "Remaining Stock": row.remaining_stock,
      "Cost Total": row.cost_total,
      Profit: row.profit,
      "Margin %": row.margin_percent
    }))),
    "Product Summary"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    sheetRows((report.category_summary || []).map((row) => ({
      Category: row.category,
      "Items Sold": row.total_quantity,
      Revenue: row.revenue,
      Discount: row.discount
    }))),
    "Category Summary"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    sheetRows((report.top_products || []).map((row) => ({
      Rank: row.rank,
      Product: row.product_name,
      "Quantity Sold": row.quantity,
      Revenue: row.revenue
    }))),
    "Top Products"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    sheetRows([
      { Method: "Cash", Amount: report.payment_summary.cash },
      { Method: "QR", Amount: report.payment_summary.qr },
      { Method: "Card", Amount: report.payment_summary.card },
      { Method: "Bank Transfer", Amount: report.payment_summary.bank_transfer },
      { Method: "E-Wallet", Amount: report.payment_summary.e_wallet },
      { Method: "Unpaid", Amount: report.payment_summary.unpaid },
      { Method: "Outstanding Amount", Amount: report.payment_summary.outstanding_amount }
    ]),
    "Payment Summary"
  );
  XLSX.utils.book_append_sheet(
    workbook,
    sheetRows((report.staff_summary || []).map((row) => ({
      Cashier: row.cashier,
      "Orders Processed": row.orders_processed,
      "Items Sold": row.items_sold,
      Revenue: row.revenue,
      "Discount Given": row.discount
    }))),
    "Staff Summary"
  );
  return XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
}

function pdfEscape(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/\(/g, "\\(")
    .replace(/\)/g, "\\)");
}

function wrapPdfLine(value, width = 92) {
  const text = String(value || "");
  if (text.length <= width) {
    return [text];
  }
  const words = text.split(/\s+/);
  const rows = [];
  let row = "";
  for (const word of words) {
    if (!row) {
      row = word;
    } else if (`${row} ${word}`.length <= width) {
      row = `${row} ${word}`;
    } else {
      rows.push(row);
      row = word;
    }
  }
  if (row) rows.push(row);
  return rows;
}

function jpegDimensions(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 4 || buffer[0] !== 0xff || buffer[1] !== 0xd8) {
    return null;
  }

  let offset = 2;
  while (offset + 9 < buffer.length) {
    if (buffer[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    const marker = buffer[offset + 1];
    const length = buffer.readUInt16BE(offset + 2);
    if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
      return {
        width: buffer.readUInt16BE(offset + 7),
        height: buffer.readUInt16BE(offset + 5)
      };
    }
    if (!length || offset + 2 + length > buffer.length) {
      break;
    }
    offset += 2 + length;
  }
  return null;
}

function pdfLogoFromSource(source) {
  const value = String(source || "").trim();
  if (!value) return null;
  try {
    let buffer = null;
    if (/^data:image\/jpe?g;base64,/i.test(value)) {
      buffer = Buffer.from(value.replace(/^data:image\/jpe?g;base64,/i, ""), "base64");
    } else {
      const logoPath = value.startsWith("/")
        ? path.join(__dirname, "public", value.replace(/^\/+/, ""))
        : path.isAbsolute(value)
          ? value
          : path.join(__dirname, "public", value);
      if (fs.existsSync(logoPath) && /\.jpe?g$/i.test(logoPath)) {
        buffer = fs.readFileSync(logoPath);
      }
    }
    const dimensions = jpegDimensions(buffer);
    if (!buffer || !dimensions) return null;
    return { buffer, ...dimensions };
  } catch (_) {
    return null;
  }
}

function buildSimplePdf(title, lines, options = {}) {
  const pageLines = [];
  for (const line of lines) {
    pageLines.push(...wrapPdfLine(line));
  }
  const maxLines = 52;
  const pages = [];
  for (let i = 0; i < pageLines.length; i += maxLines) {
    pages.push(pageLines.slice(i, i + maxLines));
  }
  if (!pages.length) {
    pages.push(["No data."]);
  }

  const objects = [];
  const addObject = (content) => {
    objects.push(content);
    return objects.length;
  };

  const catalogId = addObject("<< /Type /Catalog /Pages 2 0 R >>");
  const pagesId = addObject("");
  const fontId = addObject("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
  const logo = options.logo || null;
  const logoId = logo?.buffer
    ? addObject(Buffer.concat([
        Buffer.from(`<< /Type /XObject /Subtype /Image /Width ${logo.width} /Height ${logo.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${logo.buffer.length} >>\nstream\n`, "utf8"),
        logo.buffer,
        Buffer.from("\nendstream", "utf8")
      ]))
    : 0;
  const pageIds = [];

  for (const [index, rows] of pages.entries()) {
    const hasLogo = Boolean(logoId && index === 0);
    const logoBox = 74;
    const logoScale = hasLogo ? Math.min(logoBox / logo.width, logoBox / logo.height) : 0;
    const logoWidth = hasLogo ? Number((logo.width * logoScale).toFixed(2)) : 0;
    const logoHeight = hasLogo ? Number((logo.height * logoScale).toFixed(2)) : 0;
    const logoX = 40;
    const logoY = 800 - logoHeight + 2;
    const textX = hasLogo ? 130 : 40;
    const stream = [
      ...(hasLogo ? [
        "q",
        `${logoWidth} 0 0 ${logoHeight} ${logoX} ${logoY} cm`,
        "/Logo Do",
        "Q"
      ] : []),
      "BT",
      "/F1 10 Tf",
      "14 TL",
      `${textX} 800 Td`,
      `(${pdfEscape(index === 0 ? title : `${title} (continued)`)}) Tj`,
      "T*",
      "T*",
      ...rows.map((line) => `(${pdfEscape(line)}) Tj T*`),
      "ET"
    ].join("\n");
    const contentId = addObject(`<< /Length ${Buffer.byteLength(stream, "utf8")} >>\nstream\n${stream}\nendstream`);
    const xObjectResource = logoId ? `/XObject << /Logo ${logoId} 0 R >>` : "";
    const pageId = addObject(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${fontId} 0 R >> ${xObjectResource} >> /Contents ${contentId} 0 R >>`);
    pageIds.push(pageId);
  }

  objects[pagesId - 1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageIds.length} >>`;

  const chunks = [Buffer.from("%PDF-1.4\n", "utf8")];
  const offsets = [0];
  let byteLength = chunks[0].length;
  objects.forEach((content, index) => {
    const prefix = Buffer.from(`${index + 1} 0 obj\n`, "utf8");
    const body = Buffer.isBuffer(content) ? content : Buffer.from(String(content), "utf8");
    const suffix = Buffer.from("\nendobj\n", "utf8");
    offsets.push(byteLength);
    chunks.push(prefix, body, suffix);
    byteLength += prefix.length + body.length + suffix.length;
  });
  const xrefOffset = byteLength;
  let trailer = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1)) {
    trailer += `${String(offset).padStart(10, "0")} 00000 n \n`;
  }
  trailer += `trailer\n<< /Size ${objects.length + 1} /Root ${catalogId} 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  chunks.push(Buffer.from(trailer, "utf8"));
  return Buffer.concat(chunks);
}

function formatPdfCurrency(cents) {
  return `RM ${toMoney(cents).toFixed(2)}`;
}

function buildA4InvoicePdfBuffer(order) {
  const template = getReceiptTemplateSettings();
  const payableCents = Number(order.payable_cents ?? order.total_cents);
  const paidCents = Number(order.paid_cents || 0);
  const changeCents = Number(order.change_cents || paidCents - payableCents);
  const lines = [
    template.company_name || template.header || "Invoice",
    template.ssm_number ? `Business Registration No. (SSM): ${template.ssm_number}` : "",
    template.address || "",
    [template.phone ? `Phone: ${template.phone}` : "", template.email ? `Email: ${template.email}` : ""].filter(Boolean).join("  "),
    "",
    `Invoice Number: ${order.invoice_number || "-"}`,
    `Receipt Number: ${order.order_number || "-"}`,
    `Generated: ${order.invoice_pdf_generated_at || order.invoice_generated_at || new Date().toISOString()}`,
    `Customer: ${order.customer_name || order.member?.name || "Walk-in"}`,
    order.phone || order.member?.phone ? `Phone: ${order.phone || order.member?.phone}` : "",
    `Payment Method: ${paymentMethodLabel(order.payment_method || "")}`,
    `Payment Status: ${normalizeReceiptStatus(order.status || "paid").toUpperCase()}`,
    "",
    "Items",
    "------------------------------------------------------------"
  ].filter((line) => line !== "");

  const items = Array.isArray(order.items) ? order.items : [];
  for (const item of items) {
    const quantity = Number(item.quantity || 0);
    const unit = formatPdfCurrency(item.unit_price_cents || 0);
    const total = formatPdfCurrency(item.line_total_cents || 0);
    lines.push(`${item.product_name || "Item"} x ${quantity} @ ${unit} = ${total}`);
  }

  lines.push(
    "------------------------------------------------------------",
    `Subtotal: ${formatPdfCurrency(order.subtotal_cents || 0)}`,
    `Discount: ${formatPdfCurrency(order.discount_cents || 0)}`,
    `Tax: ${formatPdfCurrency(order.tax_cents || 0)}`,
    `Rounding: ${formatPdfCurrency(order.rounding_cents || 0)}`,
    `Total: ${formatPdfCurrency(payableCents)}`,
    `Paid: ${formatPdfCurrency(paidCents)}`,
    `Change: ${formatPdfCurrency(changeCents)}`,
    "",
    template.terms_conditions || "",
    template.thank_you_message || template.receipt_notes || "",
    template.footer || "Computer Generated Invoice"
  );

  return buildSimplePdf(order.invoice_number || "A4 Invoice", lines.filter((line) => line !== ""), {
    logo: pdfLogoFromSource(template.company_logo)
  });
}

function buildSoldItemsPdf(report) {
  const lines = [
    `Date Range: ${report.range?.start_date || "-"} to ${report.range?.end_date || "-"}`,
    `Generated: ${report.generated_at || ""}`,
    "",
    `Orders: ${report.totals?.order_count || 0}`,
    `Quantity Sold: ${report.totals?.quantity_sold || 0}`,
    `Revenue: RM ${Number(report.totals?.revenue || 0).toFixed(2)}`,
    `Discount: RM ${Number(report.totals?.discount || 0).toFixed(2)}`,
    "",
    "Top Selling Products",
    ...(report.top_products || []).map((row) => `${row.rank}. ${row.product_name} - Qty ${row.quantity}, RM ${Number(row.revenue || 0).toFixed(2)}`),
    "",
    "Payment Summary",
    `Cash: RM ${Number(report.payment_summary?.cash || 0).toFixed(2)}`,
    `QR: RM ${Number(report.payment_summary?.qr || 0).toFixed(2)}`,
    `Card: RM ${Number(report.payment_summary?.card || 0).toFixed(2)}`,
    `Bank Transfer: RM ${Number(report.payment_summary?.bank_transfer || 0).toFixed(2)}`,
    `E-Wallet: RM ${Number(report.payment_summary?.e_wallet || 0).toFixed(2)}`,
    `Unpaid: RM ${Number(report.payment_summary?.unpaid || 0).toFixed(2)}`,
    `Outstanding Amount: RM ${Number(report.payment_summary?.outstanding_amount || 0).toFixed(2)}`,
    "",
    "Sold Items",
    ...(report.item_rows || []).map((row) =>
      `${reportTime(row.created_at)} | ${row.receipt_number} | ${row.product_name} | ${row.quantity} x RM ${Number(row.unit_price || 0).toFixed(2)} | Discount RM ${Number(row.discount || 0).toFixed(2)} | Total RM ${Number(row.total_amount || 0).toFixed(2)} | ${row.cashier || "-"}`
    )
  ];
  return buildSimplePdf("Today's Sold Items Report", lines);
}

function createToken() {
  return crypto.randomBytes(24).toString("hex");
}

function getTokenFromRequest(req) {
  const bearer = String(req.headers.authorization || "");
  if (bearer.toLowerCase().startsWith("bearer ")) {
    return bearer.slice(7).trim();
  }

  const token = String(req.headers["x-auth-token"] || req.query.token || "").trim();
  return token;
}

function sessionHash(token) {
  return crypto.createHash("sha256").update(String(token || ""), "utf8").digest("hex");
}

function requestIp(req) {
  return String(req.headers["x-forwarded-for"] || req.socket?.remoteAddress || req.ip || "")
    .split(",")[0]
    .trim();
}

function requestUserAgent(req) {
  return String(req.headers["user-agent"] || "");
}

function isSuperAdmin(user) {
  return normalizeRole(user?.role) === "super_admin";
}

function isSystemOwner(user) {
  return normalizeRole(user?.role) === "system_owner";
}

function requestedRole(payload = {}) {
  return normalizeRole(payload.role || "");
}

function assertCanManageUser(req, targetUser = null, payload = {}) {
  const requested = requestedRole(payload);
  const targetRole = normalizeRole(targetUser?.role);
  if (requested === "system_owner" && !isSystemOwner(req.auth_user)) {
    const err = new Error("Only System Owner can create or assign System Owner users.");
    err.code = "FORBIDDEN";
    throw err;
  }
  if (targetUser && targetRole === "system_owner" && !isSystemOwner(req.auth_user)) {
    const err = new Error("Only System Owner can modify System Owner accounts.");
    err.code = "FORBIDDEN";
    throw err;
  }
  if (requested === "super_admin" && !(isSystemOwner(req.auth_user) || isSuperAdmin(req.auth_user))) {
    const err = new Error("Only Super Admin can create or assign Super Admin users.");
    err.code = "FORBIDDEN";
    throw err;
  }
  if (targetUser && targetRole === "super_admin" && !(isSystemOwner(req.auth_user) || isSuperAdmin(req.auth_user))) {
    const err = new Error("Only Super Admin can modify Super Admin users.");
    err.code = "FORBIDDEN";
    throw err;
  }
}

function passwordChangeRequired(user) {
  return Boolean(user?.force_password_change);
}

function passwordChangeExemptPath(req) {
  const pathOnly = String(req.path || req.originalUrl || "").split("?")[0];
  return ["/api/auth/me", "/api/auth/logout", "/api/auth/change-password"].includes(pathOnly);
}

function readSession(token) {
  if (typeof store.getSession === "function") {
    return store.getSession(sessionHash(token));
  }
  return sessions.get(token) || null;
}

function writeSession(token, session) {
  if (typeof store.saveSession === "function") {
    store.saveSession(sessionHash(token), session);
    return;
  }
  sessions.set(token, session);
}

function removeSession(token) {
  if (typeof store.deleteSession === "function") {
    store.deleteSession(sessionHash(token));
    return;
  }
  sessions.delete(token);
}

function authRequired(req, res, next) {
  const token = getTokenFromRequest(req);
  if (!token) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  const session = readSession(token);
  if (!session) {
    removeSession(token);
    return res.status(401).json({ error: "Unauthorized" });
  }

  const expiresAt = new Date(session.expires_at || 0).getTime();
  if (expiresAt && expiresAt < Date.now()) {
    removeSession(token);
    return res.status(401).json({ error: "Unauthorized" });
  }

  const user = store.getUserById(session.user_id);
  if (!user) {
    removeSession(token);
    return res.status(401).json({ error: "Unauthorized" });
  }

  writeSession(token, {
    ...session,
    last_seen_at: new Date().toISOString(),
    expires_at: new Date(Date.now() + 12 * 60 * 60 * 1000).toISOString()
  });

  req.auth_token = token;
  req.auth_user = user;
  if (passwordChangeRequired(user) && !passwordChangeExemptPath(req)) {
    return res.status(403).json({
      error: "Password change is required before accessing the system.",
      code: "PASSWORD_CHANGE_REQUIRED",
      user: mapUser(user)
    });
  }
  next();
}

function requireRole(role) {
  return (req, res, next) => {
    const actual = normalizeRole(req.auth_user?.role);
    const expected = normalizeRole(role);
    if (actual !== expected && actual !== "super_admin" && actual !== "system_owner") {
      return res.status(403).json({ error: "Forbidden" });
    }
    next();
  };
}

registerModuleApi(app, { authRequired });

app.post("/api/auth/login", (req, res) => {
  const username = String(req.body?.username || "").trim();
  const password = String(req.body?.password || "");

  if (!username || !password) {
    return res.status(400).json({ error: "Username and password are required." });
  }

  const user = store.getUserByUsername(username, { includeInactive: true });
  if (!user || !store.isPasswordValid(user, password)) {
    if (typeof store.recordLoginAttempt === "function") {
      store.recordLoginAttempt({
        username,
        ipAddress: requestIp(req),
        userAgent: requestUserAgent(req),
        status: "failed"
      });
    }
    return res.status(401).json({ error: "Invalid username or password." });
  }

  const status = String(user.status || (user.is_active ? "active" : "inactive")).toLowerCase();
  if (status !== "active") {
    if (typeof store.recordLoginAttempt === "function") {
      store.recordLoginAttempt({
        userId: user.id,
        username,
        ipAddress: requestIp(req),
        userAgent: requestUserAgent(req),
        status
      });
    }
    return res.status(403).json({ error: "This user account is inactive or suspended." });
  }

  const token = createToken();
  const tokenHash = sessionHash(token);
  writeSession(token, {
    user_id: user.id,
    username: user.username,
    role: user.role,
    created_at: new Date().toISOString(),
    last_seen_at: new Date().toISOString(),
    expires_at: new Date(Date.now() + 12 * 60 * 60 * 1000).toISOString()
  });

  if (typeof store.recordLoginAttempt === "function") {
    store.recordLoginAttempt({
      userId: user.id,
      username: user.username,
      tokenHash,
      ipAddress: requestIp(req),
      userAgent: requestUserAgent(req),
      status: "success"
    });
  }

  store.logAudit({
    action: "login",
    module: "Authentication",
    changedBy: user.username,
    description: `${user.full_name || user.username} logged in`
  });

  return res.json({ token, user: mapUser(user) });
});

app.post("/api/auth/logout", authRequired, (req, res) => {
  if (typeof store.recordLogout === "function") {
    store.recordLogout(sessionHash(req.auth_token));
  }
  store.logAudit({
    action: "logout",
    module: "Authentication",
    changedBy: req.auth_user?.username || "system",
    description: `${req.auth_user?.full_name || req.auth_user?.username || "User"} logged out`
  });
  removeSession(req.auth_token);
  res.json({ ok: true });
});

app.get("/api/auth/me", authRequired, (req, res) => {
  res.json({ user: mapUser(req.auth_user) });
});

app.post("/api/auth/change-password", authRequired, (req, res) => {
  try {
    const currentPassword = String(req.body?.current_password || req.body?.currentPassword || "");
    const password = String(req.body?.password || "");
    const confirmPassword = String(req.body?.confirm_password || req.body?.confirmPassword || "");
    const currentUser = store.getUserById(req.auth_user.id, { includeInactive: true });
    if (!currentUser || !store.isPasswordValid(currentUser, currentPassword)) {
      return res.status(400).json({ error: "Current password is incorrect." });
    }
    const user = store.resetUserPassword(
      req.auth_user.id,
      { password, confirm_password: confirmPassword, force_password_change: false },
      req.auth_user?.username || "system"
    );
    res.json({ user: mapUser(user) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to change password." });
  }
});

app.get("/api/events", authRequired, (req, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders?.();
  res.write(`event: connected\n`);
  res.write(`data: ${JSON.stringify({ ok: true, created_at: new Date().toISOString(), storage: storageInfo() })}\n\n`);
  eventClients.add(res);
  const keepAlive = setInterval(() => {
    res.write(`event: ping\n`);
    res.write(`data: ${JSON.stringify({ created_at: new Date().toISOString() })}\n\n`);
  }, 25000);
  req.on("close", () => {
    clearInterval(keepAlive);
    eventClients.delete(res);
  });
});

app.get("/api/users", authRequired, requirePermission("settings.users"), (req, res) => {
  const users = store.getUsers({ includeInactive: true }).map(mapUser);
  res.json({ users });
});

app.post("/api/users", authRequired, requirePermission("settings.users"), (req, res) => {
  try {
    assertCanManageUser(req, null, req.body || {});
    const user = store.addUser(req.body || {}, req.auth_user?.username || "admin");
    res.status(201).json({ user: mapUser(user) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create user." });
  }
});

app.get("/api/users/login-history", authRequired, requirePermission("audit.view"), (req, res) => {
  const limit = Number(req.query.limit || 500);
  const username = String(req.query.username || "").trim();
  const history = typeof store.getLoginHistory === "function" ? store.getLoginHistory({ limit, username }) : [];
  res.json({ history });
});

app.get("/api/users/:id", authRequired, requirePermission("settings.users"), (req, res) => {
  const user = store.getUserById(Number(req.params.id), { includeInactive: true });
  if (!user) {
    return res.status(404).json({ error: "User not found." });
  }
  return res.json({ user: mapUser(user) });
});

app.put("/api/users/:id", authRequired, requirePermission("settings.users"), (req, res) => {
  try {
    const existing = store.getUserById(Number(req.params.id), { includeInactive: true });
    assertCanManageUser(req, existing, req.body || {});
    const user = store.updateUser(Number(req.params.id), req.body || {}, req.auth_user?.username || "admin");
    res.json({ user: mapUser(user) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update user." });
  }
});

app.post("/api/users/:id/reset-password", authRequired, requirePermission("settings.users"), (req, res) => {
  try {
    const existing = store.getUserById(Number(req.params.id), { includeInactive: true });
    assertCanManageUser(req, existing, {});
    const user = store.resetUserPassword(
      Number(req.params.id),
      { ...(req.body || {}), force_password_change: true },
      req.auth_user?.username || "admin"
    );
    res.json({ user: mapUser(user) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to reset password." });
  }
});

app.delete("/api/users/:id", authRequired, requirePermission("settings.users"), (req, res) => {
  if (!(isSystemOwner(req.auth_user) || isSuperAdmin(req.auth_user))) {
    return res.status(403).json({ error: "Only System Owner or Super Admin can delete users." });
  }
  try {
    const existing = store.getUserById(Number(req.params.id), { includeInactive: true });
    assertCanManageUser(req, existing, {});
    const user = store.removeUser(Number(req.params.id), req.auth_user?.username || "admin");
    res.json({ user: mapUser(user) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to remove user." });
  }
});

app.get("/api/roles", authRequired, requirePermission("settings.users"), (req, res) => {
  res.json({
    roles: roleCatalog(),
    permissions: allPermissionIds,
    role_permissions: rolePermissionSnapshot(),
    modules: getModuleCatalog(req.auth_user)
  });
});

app.put("/api/roles/:role/permissions", authRequired, requirePermission("settings.roles"), (req, res) => {
  try {
    const role = normalizeRole(req.params.role);
    if (role === "system_owner" && !isSystemOwner(req.auth_user)) {
      return res.status(403).json({ error: "Only System Owner can manage System Owner permissions." });
    }
    const requested = Array.isArray(req.body?.permissions) ? req.body.permissions.map((item) => String(item || "")) : [];
    const next = {
      ...rolePermissionSnapshot(),
      [role]: requested.filter((permissionId) => allPermissionIds.includes(permissionId))
    };
    configureRolePermissions(next);
    if (typeof store.setSetting === "function") {
      store.setSetting("role_permissions_json", JSON.stringify(rolePermissionSnapshot()), req.auth_user?.username || "admin");
    }
    res.json({ roles: roleCatalog(), role_permissions: rolePermissionSnapshot() });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update role permissions." });
  }
});

app.get("/api/members", authRequired, (req, res) => {
  const search = String(req.query.search || "").trim();
  const includeInactive = String(req.query.include_inactive || "0") === "1";
  res.json({ members: store.getMembers({ search, includeInactive }) });
});

app.post("/api/members", authRequired, (req, res) => {
  try {
    const member = store.addMember(req.body || {}, req.auth_user?.username || "system");
    res.status(201).json({ member });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create member." });
  }
});

app.put("/api/members/:id", authRequired, (req, res) => {
  try {
    const member = store.updateMember(Number(req.params.id), req.body || {}, req.auth_user?.username || "system");
    res.json({ member });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update member." });
  }
});

app.delete("/api/members/:id", authRequired, (req, res) => {
  try {
    const member = store.removeMember(Number(req.params.id), req.auth_user?.username || "system");
    res.json({ member });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to remove member." });
  }
});

app.get("/api/suppliers", authRequired, requirePermission("suppliers.view"), (req, res) => {
  const search = String(req.query.search || "").trim();
  const includeInactive = String(req.query.include_inactive || "0") === "1";
  res.json({ suppliers: store.getSuppliers({ search, includeInactive }).map(mapSupplier) });
});

app.post("/api/suppliers", authRequired, requirePermission("suppliers.create"), (req, res) => {
  try {
    const supplier = store.addSupplier(req.body || {}, req.auth_user?.username || "system");
    res.status(201).json({ supplier: mapSupplier(supplier) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create supplier." });
  }
});

app.put("/api/suppliers/:id", authRequired, requirePermission("suppliers.edit"), (req, res) => {
  try {
    const supplier = store.updateSupplier(Number(req.params.id), req.body || {}, req.auth_user?.username || "system");
    res.json({ supplier: mapSupplier(supplier) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update supplier." });
  }
});

app.delete("/api/suppliers/:id", authRequired, requirePermission("suppliers.delete"), (req, res) => {
  try {
    const supplier = store.removeSupplier(Number(req.params.id), req.auth_user?.username || "system");
    res.json({ supplier: mapSupplier(supplier) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to remove supplier." });
  }
});

app.get("/api/purchasing/orders", authRequired, requirePermission("purchasing.view"), (req, res) => {
  const search = String(req.query.search || "").trim();
  const status = String(req.query.status || "").trim();
  const includeCancelled = String(req.query.include_cancelled || "0") === "1";
  const limit = Number(req.query.limit || 500);
  res.json({
    orders: store.getPurchaseOrders({ search, status, includeCancelled, limit }).map(mapPurchaseOrder)
  });
});

app.post("/api/purchasing/orders", authRequired, requirePermission("purchasing.create"), (req, res) => {
  try {
    const order = store.createPurchaseOrder(req.body || {}, req.auth_user?.username || "system");
    res.status(201).json({ order: mapPurchaseOrder(order) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create purchase order." });
  }
});

app.get("/api/purchasing/orders/:id", authRequired, requirePermission("purchasing.view"), (req, res) => {
  const order = store.getPurchaseOrderById(Number(req.params.id));
  if (!order) {
    return res.status(404).json({ error: "Purchase order not found." });
  }
  return res.json({ order: mapPurchaseOrder(order) });
});

app.post("/api/purchasing/orders/:id/receive", authRequired, requirePermission("purchasing.receive_stock"), (req, res) => {
  try {
    const order = store.receivePurchaseOrder(Number(req.params.id), req.body || {}, req.auth_user?.username || "system");
    res.json({ order: mapPurchaseOrder(order) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to receive purchase order." });
  }
});

app.get("/api/purchasing/history", authRequired, requirePermission("purchasing.view"), (req, res) => {
  const search = String(req.query.search || "").trim();
  const limit = Number(req.query.limit || 500);
  res.json({ orders: store.getPurchaseHistory({ search, limit }).map(mapPurchaseOrder) });
});

app.get("/api/finance/summary", authRequired, requirePermission("finance.receivables.view"), (req, res) => {
  const date = String(req.query.date || "").trim();
  res.json({ summary: mapFinanceSummary(store.getFinanceSummary({ date })) });
});

app.get("/api/finance/expenses", authRequired, requirePermission("finance.expenses"), (req, res) => {
  const startDate = String(req.query.start_date || "").trim();
  const endDate = String(req.query.end_date || "").trim();
  const search = String(req.query.search || "").trim();
  const limit = Number(req.query.limit || 500);
  res.json({ expenses: store.getExpenses({ start_date: startDate, end_date: endDate, search, limit }).map(mapExpense) });
});

app.post("/api/finance/expenses", authRequired, requirePermission("finance.expenses"), (req, res) => {
  try {
    const expense = store.addExpense(req.body || {}, req.auth_user?.username || "system");
    res.status(201).json({ expense: mapExpense(expense) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to record expense." });
  }
});

app.get("/api/finance/payables", authRequired, requirePermission("finance.payables.view"), (req, res) => {
  const search = String(req.query.search || "").trim();
  res.json({
    suppliers: store.getSupplierPayables({ search }).map(mapSupplier),
    payments: store.getSupplierPayments({ limit: Number(req.query.limit || 100) }).map(mapSupplierPayment)
  });
});

app.post("/api/finance/payables/:supplierId/payment", authRequired, requirePermission("purchasing.pay_supplier"), (req, res) => {
  try {
    const result = store.recordSupplierPayment(Number(req.params.supplierId), req.body || {}, req.auth_user?.username || "system");
    res.json({ supplier: mapSupplier(result.supplier), payment: mapSupplierPayment(result.payment) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to record supplier payment." });
  }
});

app.get("/api/finance/daily-closing", authRequired, requirePermission("finance.closing"), (req, res) => {
  res.json({ closings: store.getDailyClosings({ limit: Number(req.query.limit || 100) }).map(mapDailyClosing) });
});

app.post("/api/finance/daily-closing", authRequired, requirePermission("finance.closing"), (req, res) => {
  try {
    const closing = store.saveDailyClosing(req.body || {}, req.auth_user?.username || "system");
    res.status(201).json({ closing: mapDailyClosing(closing) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save daily closing." });
  }
});

app.get("/api/discounts", authRequired, (req, res) => {
  const activeNow = String(req.query.active || "0") === "1" || !isAdmin(req.auth_user);
  const includeInactive = isAdmin(req.auth_user) && String(req.query.include_inactive || "0") === "1";
  res.json({ discounts: store.getDiscounts({ includeInactive, activeNow }).map(mapDiscount) });
});

app.post("/api/discounts", authRequired, requirePermission("settings.discounts"), (req, res) => {
  try {
    const discount = mapDiscount(store.addDiscount(req.body || {}, req.auth_user?.username || "admin"));
    res.status(201).json({ discount });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create discount." });
  }
});

app.put("/api/discounts/:id", authRequired, requirePermission("settings.discounts"), (req, res) => {
  try {
    const discount = mapDiscount(store.updateDiscount(Number(req.params.id), req.body || {}, req.auth_user?.username || "admin"));
    res.json({ discount });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update discount." });
  }
});

app.delete("/api/discounts/:id", authRequired, requirePermission("settings.discounts"), (req, res) => {
  try {
    const discount = mapDiscount(store.removeDiscount(Number(req.params.id), req.auth_user?.username || "admin"));
    res.json({ discount });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to remove discount." });
  }
});

app.get("/api/discount-usage", authRequired, requirePermission("reports.view"), (req, res) => {
  const limit = Number(req.query.limit || 500);
  const usage = store.getDiscountUsage({ limit }).map((row) => ({
    ...row,
    subtotal: toMoney(row.subtotal_cents),
    discount_amount: toMoney(row.discount_cents)
  }));
  res.json({ usage });
});

app.get("/api/products", authRequired, (req, res) => {
  const search = String(req.query.search || "").trim();
  const category = String(req.query.category || "").trim();
  const subcategory = String(req.query.subcategory || "").trim();
  const brand = String(req.query.brand || "").trim();
  const supplier = String(req.query.supplier || "").trim();
  const status = String(req.query.status || "").trim();
  const stock = String(req.query.stock || "").trim();
  const sortBy = String(req.query.sort_by || "name").trim();
  const sortDir = String(req.query.sort_dir || "asc").trim();
  const includeInactive = String(req.query.include_inactive || "0") === "1";

  if (Object.hasOwn(req.query, "page") || Object.hasOwn(req.query, "page_size")) {
    const page = Number(req.query.page || 1);
    const pageSize = Number(req.query.page_size || 10);
    const paged = store.getProductsPage({
      page,
      pageSize,
      search,
      includeInactive,
      category,
      subcategory,
      brand,
      supplier,
      status,
      stock,
      sortBy,
      sortDir
    });
    return res.json({
      products: paged.rows.map((product) => mapProduct(product, req.auth_user)),
      pagination: paged.pagination,
      taxRate: TAX_RATE,
      currency: "MYR"
    });
  }

  const products = store
    .getProducts({ search, includeInactive, category, subcategory, brand, supplier, status, stock, sortBy, sortDir })
    .map((product) => mapProduct(product, req.auth_user));
  res.json({ products, taxRate: TAX_RATE, currency: "MYR" });
});

app.get("/api/product-filters", authRequired, (req, res) => {
  res.json(store.getProductFilters());
});

app.post("/api/products", authRequired, (req, res) => {
  try {
    const actor = String(req.auth_user?.username || "user");
    const body = { ...(req.body || {}) };
    if (!canViewCost(req.auth_user)) {
      delete body.cost;
      delete body.cost_price;
    }
    const product = mapProduct(store.addProduct(body, actor), req.auth_user);
    res.status(201).json({ product });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create product." });
  }
});

app.put("/api/products/:id", authRequired, (req, res) => {
  try {
    const actor = String(req.auth_user?.username || "user");
    const body = { ...(req.body || {}) };
    if (!canViewCost(req.auth_user)) {
      delete body.cost;
      delete body.cost_price;
    }
    const product = mapProduct(store.updateProduct(Number(req.params.id), body, actor), req.auth_user);
    res.json({ product });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update product." });
  }
});

app.delete("/api/products/:id", authRequired, (req, res) => {
  try {
    const actor = String(req.auth_user?.username || "user");
    const product = mapProduct(store.removeProduct(Number(req.params.id), actor), req.auth_user);
    res.json({ message: "Product removed from sale.", product });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to remove product." });
  }
});

app.get("/api/stock", authRequired, (req, res) => {
  const includeInactive = String(req.query.include_inactive || "1") === "1";
  const search = String(req.query.search || "").trim();
  const stock = store.getStock(includeInactive, search);
  res.json({ stock });
});

app.get("/api/inventory/alerts", authRequired, (req, res) => {
  const limit = Number(req.query.limit || 20);
  res.json({ alerts: store.getLowStockProducts(limit).map((product) => mapProduct(product, req.auth_user)) });
});

app.get("/api/inventory/records", authRequired, (req, res) => {
  const productId = Number(req.query.product_id || 0);
  const limit = Number(req.query.limit || 500);
  res.json({ records: store.getInventoryRecords({ productId, limit }) });
});

app.post("/api/stock/add", authRequired, (req, res) => {
  try {
    const productId = Number(req.body?.product_id);
    const quantity = Number(req.body?.quantity);
    const actor = String(req.auth_user?.username || "system");
    const product = mapProduct(store.addStock(productId, quantity, actor), req.auth_user);
    return res.json({ product });
  } catch (err) {
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to add stock." });
  }
});

app.get("/api/pending-orders", authRequired, (req, res) => {
  const actor = String(req.auth_user?.username || "");
  res.json({ orders: store.getPendingOrders({ actor }) });
});

app.post("/api/pending-orders", authRequired, (req, res) => {
  try {
    const actor = String(req.auth_user?.username || "system");
    const requestedId = Number(req.body?.id || 0);
    if (requestedId > 0) {
      const existing = store.getPendingOrders({ includeCancelled: true }).find((order) => Number(order.id) === requestedId);
      if (existing && String(existing.created_by || "").toLowerCase() !== actor.toLowerCase() && !isAdmin(req.auth_user)) {
        return res.status(403).json({ error: "You can only update your own pending orders." });
      }
    }

    const order = store.savePendingOrder(req.body || {}, actor);
    res.json({ order });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save pending order." });
  }
});

app.delete("/api/pending-orders/:id", authRequired, (req, res) => {
  try {
    const actor = String(req.auth_user?.username || "system");
    const existing = store
      .getPendingOrders({ includeCancelled: true })
      .find((order) => Number(order.id) === Number(req.params.id));
    if (existing && String(existing.created_by || "").toLowerCase() !== actor.toLowerCase() && !isAdmin(req.auth_user)) {
      return res.status(403).json({ error: "You can only cancel your own pending orders." });
    }

    const order = store.cancelPendingOrder(Number(req.params.id), actor);
    res.json({ order });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to cancel pending order." });
  }
});

app.get("/api/outstanding-orders", authRequired, (req, res) => {
  const search = String(req.query.search || "").trim();
  const filter = String(req.query.filter || "").trim();
  const orders = store.getOutstandingOrders({ search, filter }).map(mapOutstandingOrder);
  res.json({ orders });
});

app.get("/api/outstanding-summary", authRequired, (req, res) => {
  const limit = Number(req.query.limit || 5);
  const summary = store.getOutstandingSummary(limit);
  res.json({
    summary: {
      outstanding_count: summary.outstanding_count,
      todays_unpaid_sales: toMoney(summary.todays_unpaid_cents),
      outstanding_amount: toMoney(summary.outstanding_cents),
      overdue_count: summary.overdue_count,
      recent_payments: (summary.recent_payments || []).map(mapRecentPayment)
    }
  });
});

app.get("/api/customers/outstanding", authRequired, (req, res) => {
  try {
    const outstanding = store.getCustomerOutstanding({
      customer_key: String(req.query.customer_key || "").trim(),
      member_id: Number(req.query.member_id || 0),
      customer_name: String(req.query.customer_name || "").trim(),
      phone: String(req.query.phone || "").trim()
    });
    res.json({ outstanding: mapCustomerOutstanding(outstanding) });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to load customer outstanding balance." });
  }
});

app.get("/api/customer-ledger", authRequired, (req, res) => {
  const search = String(req.query.search || "").trim();
  const customerKey = String(req.query.customer_key || "").trim();
  const outstandingReport = store.getCustomerOutstandingReport({ search }).map(mapCustomerOutstandingReportRow);
  const ledger = store.getCustomerLedgerReport({ search, customer_key: customerKey }).map(mapCustomerLedgerRow);
  const summary = mapCustomerLedgerSummary(store.getCustomerLedgerSummary({ search }));
  res.json({ outstanding_report: outstandingReport, ledger, summary });
});

app.get("/api/reports/customer-ledger", authRequired, (req, res) => {
  const search = String(req.query.search || "").trim();
  const customerKey = String(req.query.customer_key || "").trim();
  const outstandingReport = store.getCustomerOutstandingReport({ search }).map(mapCustomerOutstandingReportRow);
  const ledger = store.getCustomerLedgerReport({ search, customer_key: customerKey }).map(mapCustomerLedgerRow);
  const summary = mapCustomerLedgerSummary(store.getCustomerLedgerSummary({ search }));
  res.json({ outstanding_report: outstandingReport, ledger, summary });
});

app.post("/api/orders", authRequired, (req, res) => {
  try {
    const requestedDiscount = req.body?.discount_id
      ? { id: req.body.discount_id }
      : req.body?.discount
        ? req.body.discount
        : null;
    if (!isAdmin(req.auth_user) && !isApprovedStaffDiscount(requestedDiscount)) {
      return res.status(403).json({ error: "Staff can only apply approved discounts." });
    }

    const order = store.createOrder(
      {
        ...(req.body || {}),
        user_id: req.auth_user.id
      },
      TAX_RATE
    );

    const fullOrder = store.getOrderById(order.id) || order;
    const payableCents = Number(fullOrder.payable_cents ?? fullOrder.total_cents);
    const paidCents = Number(fullOrder.paid_cents ?? order.paid_cents ?? payableCents);
    const outstandingCents = Math.max(0, payableCents - paidCents);
    const items = (fullOrder.items || order.items || []).map((item) => ({
      ...item,
      unit_price: toMoney(item.unit_price_cents),
      line_total: toMoney(item.line_total_cents)
    }));
    res.status(201).json({
      order: {
        ...fullOrder,
        invoice_number: String(fullOrder.invoice_number || ""),
        subtotal: toMoney(fullOrder.subtotal_cents),
        tax: toMoney(fullOrder.tax_cents),
        discount: toMoney(fullOrder.discount_cents),
        total: toMoney(payableCents),
        raw_total: toMoney(fullOrder.total_cents),
        rounding: toMoney(fullOrder.rounding_cents),
        paid: toMoney(paidCents),
        change: toMoney(paidCents - payableCents),
        outstanding_balance: toMoney(outstandingCents),
        payment_status: normalizeReceiptStatus(fullOrder.status || "paid"),
        payment_method: fullOrder.payment_method || "",
        user: fullOrder.user || null,
        member: fullOrder.member || null,
        items
      }
    });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to create order." });
  }
});

app.post("/api/orders/:id/collect-payment", authRequired, (req, res) => {
  try {
    const updated = store.collectOutstandingPayment(Number(req.params.id), req.body || {}, req.auth_user?.username || "system");
    const payableCents = Number(updated.payable_cents ?? updated.total_cents);
    const paidCents = Number(updated.paid_cents || 0);
    const outstandingCents = Math.max(0, payableCents - paidCents);
    const items = (updated.items || []).map((item) => ({
      ...item,
      unit_price: toMoney(item.unit_price_cents),
      line_total: toMoney(item.line_total_cents)
    }));

    return res.json({
      order: {
        ...updated,
        invoice_number: String(updated.invoice_number || ""),
        subtotal: toMoney(updated.subtotal_cents),
        tax: toMoney(updated.tax_cents),
        discount: toMoney(updated.discount_cents),
        total: toMoney(payableCents),
        raw_total: toMoney(updated.total_cents),
        rounding: toMoney(updated.rounding_cents),
        paid: toMoney(paidCents),
        change: toMoney(paidCents - payableCents),
        outstanding_balance: toMoney(outstandingCents),
        payment_status: normalizeReceiptStatus(updated.status || "paid"),
        payment_method: updated.payment_method || "",
        user: updated.user || null,
        member: updated.member || null,
        items
      }
    });
  } catch (err) {
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to collect payment." });
  }
});

app.post("/api/orders/:id/clear-outstanding", authRequired, (req, res) => {
  try {
    const updated = store.clearOutstandingInvoice(Number(req.params.id), req.auth_user?.username || "system");
    return res.json({
      message: "Outstanding payment has been fully settled.",
      order: mapOrderForApi(updated)
    });
  } catch (err) {
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to clear outstanding invoice." });
  }
});

app.get("/api/orders/:id", authRequired, (req, res) => {
  const order = store.getOrderById(Number(req.params.id));
  if (!order) {
    return res.status(404).json({ error: "Order not found." });
  }

  const items = order.items.map((item) => ({
    ...item,
    unit_price: toMoney(item.unit_price_cents),
    line_total: toMoney(item.line_total_cents)
  }));

  const payableCents = Number(order.payable_cents ?? order.total_cents);
  const paidCents = Number(order.paid_cents || 0);
  const outstandingCents = Math.max(0, payableCents - paidCents);
  return res.json({
    order: {
      ...order,
      invoice_number: String(order.invoice_number || ""),
      subtotal: toMoney(order.subtotal_cents),
      tax: toMoney(order.tax_cents),
      discount: toMoney(order.discount_cents),
      total: toMoney(payableCents),
      raw_total: toMoney(order.total_cents),
      rounding: toMoney(order.rounding_cents),
      paid: toMoney(paidCents),
      change: toMoney(order.change_cents),
      outstanding_balance: toMoney(outstandingCents),
      payment_status: normalizeReceiptStatus(order.status || "paid"),
      payment_method: order.payment_method || "",
      user: order.user || null,
      member: order.member || null,
      items
    }
  });
});

app.put("/api/orders/:id", authRequired, (req, res) => {
  if (!isAdmin(req.auth_user) && !staffCanUpdateReceipt(req.body || {})) {
    return res.status(403).json({ error: "Staff can only update authorised payment statuses." });
  }

  try {
    const updated = store.updateOrder(Number(req.params.id), req.body || {}, req.auth_user?.username || "system");
    const payableCents = Number(updated.payable_cents ?? updated.total_cents);
    const paidCents = Number(updated.paid_cents || 0);
    const outstandingCents = Math.max(0, payableCents - paidCents);
    const items = (updated.items || []).map((item) => ({
      ...item,
      unit_price: toMoney(item.unit_price_cents),
      line_total: toMoney(item.line_total_cents)
    }));

    return res.json({
      order: {
        ...updated,
        invoice_number: String(updated.invoice_number || ""),
        subtotal: toMoney(updated.subtotal_cents),
        tax: toMoney(updated.tax_cents),
        discount: toMoney(updated.discount_cents),
        total: toMoney(payableCents),
        raw_total: toMoney(updated.total_cents),
        rounding: toMoney(updated.rounding_cents),
        paid: toMoney(paidCents),
        change: toMoney(updated.change_cents),
        outstanding_balance: toMoney(outstandingCents),
        payment_status: normalizeReceiptStatus(updated.status || "paid"),
        payment_method: updated.payment_method || "",
        user: updated.user || null,
        member: updated.member || null,
        items
      }
    });
  } catch (err) {
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to update order." });
  }
});

app.delete("/api/orders/:id", authRequired, requirePermission("receipts.delete"), (req, res) => {
  try {
    const deleted = store.deleteOrder(Number(req.params.id), req.auth_user?.username || "admin");
    return res.json({ message: "Receipt deleted.", order: deleted });
  } catch (err) {
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to delete receipt." });
  }
});

app.get("/api/orders", authRequired, (req, res) => {
  const search = String(req.query.search || "").trim();
  const status = String(req.query.status || "").trim();
  const hasPaging = Object.hasOwn(req.query, "page") || Object.hasOwn(req.query, "page_size");
  if (hasPaging) {
    const page = Number(req.query.page || 1);
    const pageSize = Number(req.query.page_size || 10);
    const paged = store.getOrdersPage({ page, pageSize, search, status });
    const orders = paged.rows.map((order) => ({
      ...order,
      invoice_number: String(order.invoice_number || ""),
      subtotal: toMoney(order.subtotal_cents),
      discount: toMoney(order.discount_cents),
      tax: toMoney(order.tax_cents),
      total: toMoney(order.payable_cents ?? order.total_cents),
      raw_total: toMoney(order.total_cents),
      rounding: toMoney(order.rounding_cents),
      paid: toMoney(order.paid_cents),
      outstanding_balance: toMoney(order.outstanding_cents),
      payment_status: normalizeReceiptStatus(order.status || "paid")
    }));
    return res.json({ orders, pagination: paged.pagination });
  }

  const limit = Number(req.query.limit || 200);
  const orders = store.getOrders(limit, search, status).map((order) => ({
    ...order,
    invoice_number: String(order.invoice_number || ""),
    subtotal: toMoney(order.subtotal_cents),
    discount: toMoney(order.discount_cents),
    tax: toMoney(order.tax_cents),
    total: toMoney(order.payable_cents ?? order.total_cents),
    raw_total: toMoney(order.total_cents),
    rounding: toMoney(order.rounding_cents),
    paid: toMoney(order.paid_cents),
    outstanding_balance: toMoney(order.outstanding_cents),
    payment_status: normalizeReceiptStatus(order.status || "paid")
  }));
  return res.json({ orders });
});

app.post("/api/orders/:id/a4-invoice", authRequired, requirePermission("receipts.print"), (req, res) => {
  try {
    const orderId = Number(req.params.id);
    const order = store.getOrderById(orderId);
    if (!order) {
      return res.status(404).json({ error: "Order not found." });
    }

    const generatedAt = new Date().toISOString();
    const invoiceSettings = getInvoiceSettings();
    const invoiceOrder = store.ensureOrderInvoiceNumber(
      order.id,
      invoiceSettings,
      order.invoice_generated_at || generatedAt,
      req.auth_user?.username || "system"
    );
    const invoiceNumber = String(invoiceOrder.invoice_number || "").trim();
    if (!invoiceNumber) {
      throw new Error("Failed to assign invoice number.");
    }
    const pdfSettings = getPdfExportSettings();
    if (pdfSettings.folder === "custom" && !path.isAbsolute(pdfSettings.custom_path)) {
      return res.status(400).json({ error: "Configured custom PDF export folder is invalid." });
    }
    const resolvedFolder = resolvePdfExportFolder(pdfSettings);
    fs.mkdirSync(resolvedFolder.folder_path, { recursive: true });
    const target = uniquePdfTarget(resolvedFolder.folder_path, invoiceNumber);
    const pdfOrder = {
      ...invoiceOrder,
      invoice_number: invoiceNumber,
      invoice_generated_at: invoiceOrder.invoice_generated_at || generatedAt,
      invoice_pdf_generated_at: generatedAt
    };
    const buffer = buildA4InvoicePdfBuffer(pdfOrder);
    fs.writeFileSync(target.file_path, buffer);

    const updated = store.recordInvoiceExport(
      order.id,
      {
        invoice_settings: invoiceSettings,
        invoice_number: invoiceNumber,
        pdf_file_name: target.file_name,
        pdf_file_path: target.file_path,
        generated_at: generatedAt,
        print: Boolean(req.body?.print)
      },
      req.auth_user?.username || "system"
    );

    const fullOrder = updated || store.getOrderById(order.id) || pdfOrder;
    const payableCents = Number(fullOrder.payable_cents ?? fullOrder.total_cents);
    const paidCents = Number(fullOrder.paid_cents || 0);
    const items = (fullOrder.items || []).map((item) => ({
      ...item,
      unit_price: toMoney(item.unit_price_cents),
      line_total: toMoney(item.line_total_cents)
    }));

    return res.json({
      invoice: {
        invoice_number: fullOrder.invoice_number || invoiceNumber,
        file_name: fullOrder.invoice_pdf_file_name || target.file_name,
        file_path: fullOrder.invoice_pdf_file_path || target.file_path,
        save_location: resolvedFolder.folder_path,
        generated_at: fullOrder.invoice_pdf_generated_at || generatedAt,
        print_count: Number(fullOrder.invoice_print_count || 0),
        export_count: Number(fullOrder.invoice_export_count || 0)
      },
      order: {
        ...fullOrder,
        invoice_number: String(fullOrder.invoice_number || invoiceNumber),
        subtotal: toMoney(fullOrder.subtotal_cents),
        tax: toMoney(fullOrder.tax_cents),
        discount: toMoney(fullOrder.discount_cents),
        total: toMoney(payableCents),
        raw_total: toMoney(fullOrder.total_cents),
        rounding: toMoney(fullOrder.rounding_cents),
        paid: toMoney(paidCents),
        change: toMoney(fullOrder.change_cents),
        outstanding_balance: toMoney(Math.max(0, payableCents - paidCents)),
        payment_status: normalizeReceiptStatus(fullOrder.status || "paid"),
        payment_method: fullOrder.payment_method || "",
        user: fullOrder.user || null,
        member: fullOrder.member || null,
        items
      }
    });
  } catch (err) {
    console.error("A4 invoice export failed", err);
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to export A4 invoice PDF." });
  }
});

app.get("/api/orders/:id/invoice-pdf", authRequired, requirePermission("receipts.view"), (req, res) => {
  const order = store.getOrderById(Number(req.params.id));
  if (!order || !order.invoice_pdf_file_path) {
    return res.status(404).json({ error: "Invoice PDF not found." });
  }
  const filePath = path.resolve(order.invoice_pdf_file_path);
  if (!fs.existsSync(filePath) || !filePath.toLowerCase().endsWith(".pdf")) {
    return res.status(404).json({ error: "Invoice PDF file is missing." });
  }
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `inline; filename=\"${path.basename(filePath).replace(/"/g, "")}\"`);
  return res.sendFile(filePath);
});

app.get("/api/edit-logs", authRequired, requirePermission("audit.view"), (req, res) => {
  const hasPaging = Object.hasOwn(req.query, "page") || Object.hasOwn(req.query, "page_size");
  if (hasPaging) {
    const page = Number(req.query.page || 1);
    const pageSize = Number(req.query.page_size || 10);
    const paged = store.getEditLogsPage({ page, pageSize });
    const logs = paged.rows.map((log) => ({
      ...log,
      details_json: parseLogDetails(log.details)
    }));
    return res.json({ logs, pagination: paged.pagination });
  }

  const limit = Number(req.query.limit || 500);
  const logs = store.getEditLogs(limit).map((log) => ({
    ...log,
    details_json: parseLogDetails(log.details)
  }));
  return res.json({ logs });
});

app.get("/api/admin/summary", authRequired, requirePermission("dashboard.view_profit"), (req, res) => {
  const targetDate = String(req.query.date || "").trim();
  const summary = store.getTodaySummary(targetDate);
  const paymentMethods = {};
  for (const [method, cents] of Object.entries(summary.payment_methods_cents || {})) {
    paymentMethods[method] = toMoney(cents);
  }

  res.json({
    summary: {
      date: summary.date,
      order_count: summary.order_count,
      customer_count: summary.customer_count,
      subtotal: toMoney(summary.subtotal_cents),
      total: toMoney(summary.total_cents),
      cogs: toMoney(summary.cogs_cents),
      profit: toMoney(summary.profit_cents),
      rounding: toMoney(summary.rounding_cents),
      payment_methods: paymentMethods
    }
  });
});

app.get("/api/sales/summary", authRequired, requirePermission("reports.view"), (req, res) => {
  const targetDate = String(req.query.date || "").trim();
  const limit = Number(req.query.limit || 20);
  const summary = store.getSalesBreakdown(targetDate, limit);
  res.json({ summary });
});

app.get("/api/reports/todays-sold-items", authRequired, (req, res) => {
  const rawReport = store.getSoldItemsReport(soldItemsQuery(req));
  const report = mapSoldItemsReport(rawReport, isAdmin(req.auth_user));
  res.json({ report });
});

app.get("/api/reports/todays-sold-items/export.xlsx", authRequired, requirePermission("reports.export"), (req, res) => {
  const rawReport = store.getSoldItemsReport(soldItemsQuery(req));
  const report = mapSoldItemsReport(rawReport, true);
  const buffer = buildSoldItemsWorkbook(report);
  const fileDate = String(report.range?.start_date || "report").replace(/[^0-9-]/g, "");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=\"todays-sold-items-${fileDate}.xlsx\"`);
  res.send(buffer);
});

app.get("/api/reports/todays-sold-items/export.pdf", authRequired, requirePermission("reports.export"), (req, res) => {
  const rawReport = store.getSoldItemsReport(soldItemsQuery(req));
  const report = mapSoldItemsReport(rawReport, true);
  const buffer = buildSoldItemsPdf(report);
  const fileDate = String(report.range?.start_date || "report").replace(/[^0-9-]/g, "");
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename=\"todays-sold-items-${fileDate}.pdf\"`);
  res.send(buffer);
});

app.get("/api/backups", authRequired, requirePermission("settings.backup"), (req, res) => {
  if (typeof store.listBackups !== "function") {
    return res.status(400).json({ error: "Database backup is available only in database storage mode." });
  }
  res.json({ backups: store.listBackups() });
});

app.post("/api/backups", authRequired, requirePermission("settings.backup"), (req, res) => {
  if (typeof store.createBackup !== "function") {
    return res.status(400).json({ error: "Database backup is available only in database storage mode." });
  }
  try {
    const backup = store.createBackup(req.auth_user?.username || "admin");
    return res.status(201).json({ backup });
  } catch (err) {
    return res.status(500).json({ error: err.message || "Failed to create database backup." });
  }
});

app.post("/api/backups/restore", authRequired, requirePermission("settings.backup"), (req, res) => {
  if (typeof store.restoreBackup !== "function") {
    return res.status(400).json({ error: "Database restore is available only in database storage mode." });
  }
  try {
    const restored = store.restoreBackup(req.body?.file_name, req.auth_user?.username || "admin");
    broadcastDataChange("database_restored", restored);
    return res.json({ restored });
  } catch (err) {
    return res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to restore database backup." });
  }
});

app.get("/api/settings", authRequired, requirePermission("settings.view"), (req, res) => {
  const companyProfile = getCompanyProfileSettings();
  res.json({
    settings: {
      business_name: companyProfile.company_name,
      currency: "MYR",
      tax_rate: TAX_RATE,
      data_file: dataFile,
      storage: storageHealth(),
      backup_dir: backupDir,
      receipt_statuses: ["paid", "unpaid", "partially_paid", "refunded", "void", "cancelled"],
      payment_methods: ["cash", "duitnow_qr", "credit_card", "debit_card", "bank_transfer", "e_wallet", "havent_paid"],
      roles: roleCatalog(),
      role_permissions: rolePermissionSnapshot(),
      modules: getModuleCatalog(req.auth_user),
      permissions: permissionsForSettings(),
      network: getNetworkAccessInfo(),
      company_profile: companyProfile,
      receipt_template: getReceiptTemplateSettings(),
      printer_settings: getPrinterSettings(),
      invoice_settings: getInvoiceSettings(),
      last_invoice_number: store.getSetting?.("last_invoice_number", "") || "",
      pdf_export_settings: getPdfExportSettings(),
      pdf_export_resolved_folder: resolvePdfExportFolder()
    }
  });
});

app.get("/api/company-profile", authRequired, (req, res) => {
  res.json({ company_profile: getCompanyProfileSettings() });
});

app.put("/api/company-profile", authRequired, requirePermission("settings.company"), (req, res) => {
  try {
    if (typeof store.updateCompanySettings !== "function") {
      return res.status(500).json({ error: "Company settings storage is unavailable." });
    }
    const companyProfile = store.updateCompanySettings(req.body || {}, req.auth_user?.username || "admin");
    res.json({ company_profile: companyProfile });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save company information." });
  }
});

app.get("/api/receipt-template", authRequired, (req, res) => {
  res.json({ template: getReceiptTemplateSettings() });
});

app.put("/api/receipt-template", authRequired, requirePermission("settings.printers"), (req, res) => {
  try {
    const template = normalizeTemplatePayload(req.body || {});
    store.setSetting("receipt_template_json", JSON.stringify(template), req.auth_user?.username || "admin");
    res.json({ template });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save receipt template." });
  }
});

app.get("/api/printer-settings", authRequired, (req, res) => {
  res.json({ printer_settings: getPrinterSettings() });
});

app.put("/api/printer-settings", authRequired, requirePermission("settings.printers"), (req, res) => {
  try {
    const printerSettings = normalizePrinterPayload(req.body || {});
    store.setSetting("printer_settings_json", JSON.stringify(printerSettings), req.auth_user?.username || "admin");
    res.json({ printer_settings: printerSettings });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save printer settings." });
  }
});

app.get("/api/invoice-settings", authRequired, requirePermission("settings.view"), (req, res) => {
  res.json({
    invoice_settings: getInvoiceSettings(),
    last_invoice_number: store.getSetting?.("last_invoice_number", "") || ""
  });
});

app.put("/api/invoice-settings", authRequired, requirePermission("settings.printers"), (req, res) => {
  if (!(isSystemOwner(req.auth_user) || isSuperAdmin(req.auth_user))) {
    return res.status(403).json({ error: "Only System Owner or Super Admin can change invoice numbering settings." });
  }
  try {
    const invoiceSettings = normalizeInvoiceSettingsPayload(req.body || {});
    store.setSetting("invoice_settings_json", JSON.stringify(invoiceSettings), req.auth_user?.username || "admin");
    res.json({
      invoice_settings: invoiceSettings,
      last_invoice_number: store.getSetting?.("last_invoice_number", "") || ""
    });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save invoice settings." });
  }
});

app.get("/api/pdf-export-settings", authRequired, requirePermission("settings.view"), (req, res) => {
  res.json({
    pdf_export_settings: getPdfExportSettings(),
    resolved_folder: resolvePdfExportFolder()
  });
});

app.put("/api/pdf-export-settings", authRequired, requirePermission("settings.printers"), (req, res) => {
  try {
    const pdfExportSettings = normalizePdfExportPayload(req.body || {});
    if (pdfExportSettings.folder === "custom" && !path.isAbsolute(pdfExportSettings.custom_path)) {
      return res.status(400).json({ error: "Custom PDF export folder must be an absolute path." });
    }
    const resolved = resolvePdfExportFolder(pdfExportSettings);
    fs.mkdirSync(resolved.folder_path, { recursive: true });
    store.setSetting("pdf_export_settings_json", JSON.stringify(pdfExportSettings), req.auth_user?.username || "admin");
    res.json({ pdf_export_settings: pdfExportSettings, resolved_folder: resolved });
  } catch (err) {
    res.status(statusFromCode(err.code)).json({ error: err.message || "Failed to save PDF export settings." });
  }
});

app.get("/api/network/access", async (_, res) => {
  try {
    const access = getNetworkAccessInfo();
    const qr_code = await QRCode.toDataURL(access.login_url, {
      errorCorrectionLevel: "M",
      margin: 1,
      width: 280
    });

    res.json({ access: { ...access, qr_code } });
  } catch (err) {
    res.status(500).json({ error: err.message || "Failed to generate mobile access details." });
  }
});

app.get("/api/health", (_, res) => {
  const database = storageHealth();
  res.status(database.ok ? 200 : 503).json({ ok: database.ok, database, legacy_excel_file: dataFile, network: getNetworkAccessInfo() });
});

function scheduleAutomaticBackups() {
  if (typeof store.createBackup !== "function") {
    return;
  }
  const minutes = Math.max(15, Number(process.env.POS_BACKUP_INTERVAL_MINUTES || 1440));
  setInterval(() => {
    try {
      store.createBackup("automatic");
    } catch (err) {
      console.error("Automatic database backup failed:", err.message || err);
    }
  }, minutes * 60 * 1000).unref?.();
}

scheduleAutomaticBackups();

app.listen(PORT, HOST, () => {
  const access = getNetworkAccessInfo();
  console.log(`Tunas Badminton Sport Centre server running on http://127.0.0.1:${PORT}`);
  if (access.ip_address) {
    console.log(`Mobile POS access available on ${access.login_url}`);
  } else {
    console.log("Mobile POS access unavailable: no LAN IPv4 address detected.");
  }
});
