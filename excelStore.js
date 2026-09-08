const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const crypto = require("crypto");

const HEADERS = {
  products: [
    "id",
    "sku",
    "name",
    "category",
    "subcategory",
    "barcode",
    "brand",
    "supplier",
    "description",
    "template_type",
    "attributes_json",
    "unit_price_cents",
    "cost_cents",
    "quantity_on_hand",
    "reorder_level",
    "is_active",
    "created_at",
    "updated_at"
  ],
  orders: [
    "id",
    "order_number",
    "invoice_number",
    "invoice_generated_at",
    "invoice_pdf_file_name",
    "invoice_pdf_file_path",
    "invoice_pdf_generated_at",
    "invoice_print_count",
    "invoice_export_count",
    "invoice_last_printed_at",
    "invoice_last_exported_at",
    "user_id",
    "member_id",
    "customer_name",
    "phone",
    "status",
    "subtotal_cents",
    "tax_cents",
    "discount_cents",
    "total_cents",
    "rounding_cents",
    "payable_cents",
    "discount_label",
    "notes",
    "created_at"
  ],
  order_items: [
    "id",
    "order_id",
    "product_id",
    "product_name",
    "quantity",
    "unit_price_cents",
    "unit_cost_cents",
    "line_total_cents"
  ],
  payments: ["id", "order_id", "payment_method", "amount_cents", "notes", "paid_at"],
  pending_orders: [
    "id",
    "label",
    "customer_name",
    "phone",
    "member_id",
    "notes",
    "discount_id",
    "items_json",
    "status",
    "created_by",
    "created_at",
    "updated_at"
  ],
  edit_logs: ["id", "product_id", "action", "changed_by", "details", "created_at"],
  users: [
    "id",
    "username",
    "password",
    "full_name",
    "role",
    "email",
    "phone",
    "employee_id",
    "profile_photo",
    "status",
    "is_active",
    "force_password_change",
    "last_login_at",
    "created_at",
    "updated_at"
  ],
  login_history: [
    "id",
    "user_id",
    "username",
    "login_time",
    "logout_time",
    "ip_address",
    "device_type",
    "browser",
    "login_status",
    "session_token_hash"
  ],
  members: [
    "id",
    "member_no",
    "name",
    "phone",
    "email",
    "tier",
    "discount_percent",
    "is_active",
    "created_at",
    "updated_at"
  ],
  suppliers: [
    "id",
    "name",
    "contact_person",
    "phone",
    "email",
    "address",
    "payment_terms",
    "notes",
    "outstanding_cents",
    "is_active",
    "created_at",
    "updated_at"
  ],
  purchase_orders: [
    "id",
    "po_number",
    "supplier_id",
    "supplier_name",
    "status",
    "order_date",
    "expected_date",
    "received_at",
    "subtotal_cents",
    "discount_cents",
    "total_cents",
    "notes",
    "created_by",
    "created_at",
    "updated_at"
  ],
  purchase_order_items: [
    "id",
    "purchase_order_id",
    "product_id",
    "product_name",
    "sku",
    "quantity_ordered",
    "quantity_received",
    "unit_cost_cents",
    "line_total_cents"
  ],
  expenses: [
    "id",
    "expense_date",
    "category",
    "vendor",
    "description",
    "payment_method",
    "amount_cents",
    "created_by",
    "created_at",
    "updated_at"
  ],
  daily_closings: [
    "id",
    "closing_date",
    "opening_cash_cents",
    "cash_sales_cents",
    "cash_expenses_cents",
    "expected_cash_cents",
    "counted_cash_cents",
    "variance_cents",
    "notes",
    "created_by",
    "created_at",
    "updated_at"
  ],
  supplier_payments: [
    "id",
    "supplier_id",
    "supplier_name",
    "payment_method",
    "amount_cents",
    "notes",
    "paid_at",
    "created_by"
  ],
  discounts: [
    "id",
    "name",
    "code",
    "type",
    "scope",
    "value",
    "start_date",
    "end_date",
    "min_purchase_cents",
    "max_discount_cents",
    "applicable_categories",
    "applicable_product_ids",
    "applicable_member_tiers",
    "member_tier",
    "promo_type",
    "promo_buy_qty",
    "promo_free_qty",
    "promo_discount_percent",
    "promo_spend_cents",
    "promo_discount_cents",
    "is_active",
    "created_by",
    "created_at",
    "updated_at"
  ],
  discount_usage: [
    "id",
    "discount_id",
    "discount_name",
    "discount_code",
    "order_id",
    "order_number",
    "user_id",
    "staff_name",
    "member_id",
    "subtotal_cents",
    "discount_cents",
    "created_at"
  ],
  inventory_records: [
    "id",
    "product_id",
    "action",
    "quantity_delta",
    "before_qty",
    "after_qty",
    "reason",
    "changed_by",
    "created_at"
  ],
  customers: ["id", "member_no", "name", "phone", "email", "tier", "discount_percent", "is_active", "created_at", "updated_at"],
  sales: [
    "id",
    "order_number",
    "invoice_number",
    "customer_name",
    "phone",
    "status",
    "subtotal_cents",
    "discount_cents",
    "tax_cents",
    "payable_cents",
    "paid_cents",
    "outstanding_cents",
    "payment_method",
    "cashier",
    "notes",
    "created_at"
  ],
  outstanding_invoices: [
    "id",
    "order_id",
    "receipt_number",
    "customer_key",
    "customer_name",
    "phone",
    "member_id",
    "description",
    "original_cents",
    "paid_cents",
    "balance_cents",
    "status",
    "created_at",
    "updated_at"
  ],
  customer_ledger: [
    "id",
    "customer_key",
    "customer_name",
    "phone",
    "member_id",
    "order_id",
    "receipt_number",
    "description",
    "debit_cents",
    "credit_cents",
    "created_at",
    "created_by"
  ],
  company_settings: [
    "id",
    "company_logo",
    "company_name",
    "business_registration_number",
    "registration_date",
    "state",
    "nature_of_business_json",
    "business_address",
    "phone",
    "email",
    "website",
    "sst_registration_number",
    "created_at",
    "updated_at"
  ],
  settings: ["key", "value", "updated_at"],
  meta: ["key", "value"]
};

function nowISO() {
  return new Date().toISOString();
}

function toInt(value, fallback = 0) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return fallback;
  }
  return Math.trunc(numeric);
}

function toNonNegativeInt(value, fallback = 0) {
  const numeric = toInt(value, fallback);
  return numeric < 0 ? fallback : numeric;
}

function centsFromValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 0) {
    return 0;
  }
  return Math.round(numeric * 100);
}

function orderNumber(sequenceValue, value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  const datePart = dateKeyLocal(date).replace(/-/g, "");
  const sequence = String(toNonNegativeInt(sequenceValue, 0)).padStart(5, "0");
  return `ORD-${datePart}-${sequence}`;
}

function safeJsonParse(value, fallback) {
  try {
    return JSON.parse(String(value || ""));
  } catch (_) {
    return fallback;
  }
}

function normalizeInvoiceSettings(settings = {}) {
  const dateFormat = String(settings.date_format || settings.number_format || "YYMM").trim().toUpperCase();
  const reset = String(settings.reset || settings.reset_numbering || "monthly").trim().toLowerCase();
  const numberLength = Math.min(8, Math.max(3, toNonNegativeInt(settings.number_length, 3)));
  const prefix = String(settings.prefix || settings.invoice_prefix || "Invoice")
    .trim()
    .replace(/\s+/g, "-")
    .replace(/[^A-Za-z0-9_-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "") || "Invoice";
  return {
    prefix,
    date_format: ["YYMM", "DDMM", "YYYYMM", "YYYY", "NONE"].includes(dateFormat) ? dateFormat : "YYMM",
    number_length: numberLength,
    reset: ["daily", "monthly", "yearly", "never"].includes(reset) ? reset : "monthly"
  };
}

function invoiceDateParts(value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  const year = String(date.getFullYear());
  const yy = year.slice(-2);
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return { year, yy, month, day };
}

function invoiceDateToken(value = new Date(), format = "YYMM") {
  const parts = invoiceDateParts(value);
  if (format === "DDMM") return `${parts.day}${parts.month}`;
  if (format === "YYYYMM") return `${parts.year}${parts.month}`;
  if (format === "YYYY") return parts.year;
  if (format === "NONE") return "";
  return `${parts.yy}${parts.month}`;
}

function invoicePeriodKey(value = new Date(), reset = "monthly") {
  const parts = invoiceDateParts(value);
  if (reset === "daily") return `${parts.year}${parts.month}${parts.day}`;
  if (reset === "yearly") return parts.year;
  if (reset === "never") return "never";
  return `${parts.year}${parts.month}`;
}

function invoiceNumberFromIdentity(settings, token, sequence) {
  const padded = String(toNonNegativeInt(sequence, 0)).padStart(settings.number_length, "0");
  return token ? `${settings.prefix}-${token}-${padded}` : `${settings.prefix}-${padded}`;
}

function purchaseOrderNumber(sequenceValue, value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  const datePart = dateKeyLocal(date).replace(/-/g, "");
  const sequence = String(toNonNegativeInt(sequenceValue, 0)).padStart(5, "0");
  return `PO-${datePart}-${sequence}`;
}

function roundToNearest5Cents(cents) {
  return Math.round(Number(cents || 0) / 5) * 5;
}

function dateKeyLocal(value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDateKey(value) {
  const match = String(value || "").trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function addDays(value, days) {
  const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  date.setDate(date.getDate() + Number(days || 0));
  return date;
}

function reportDateRange({ preset = "today", startDate = "", endDate = "" } = {}) {
  const today = new Date();
  const normalizedPreset = String(preset || "today").trim().toLowerCase().replace(/[\s-]+/g, "_");

  if (normalizedPreset === "yesterday") {
    const yesterday = addDays(today, -1);
    const key = dateKeyLocal(yesterday);
    return { preset: "yesterday", start_date: key, end_date: key };
  }

  if (normalizedPreset === "this_week") {
    const day = today.getDay();
    const daysFromMonday = day === 0 ? 6 : day - 1;
    return {
      preset: "this_week",
      start_date: dateKeyLocal(addDays(today, -daysFromMonday)),
      end_date: dateKeyLocal(today)
    };
  }

  if (normalizedPreset === "this_month") {
    return {
      preset: "this_month",
      start_date: dateKeyLocal(new Date(today.getFullYear(), today.getMonth(), 1)),
      end_date: dateKeyLocal(today)
    };
  }

  if (normalizedPreset === "custom") {
    const start = parseDateKey(startDate);
    const end = parseDateKey(endDate);
    if (start && end) {
      return {
        preset: "custom",
        start_date: dateKeyLocal(start <= end ? start : end),
        end_date: dateKeyLocal(start <= end ? end : start)
      };
    }
  }

  const key = dateKeyLocal(today);
  return { preset: "today", start_date: key, end_date: key };
}

const PAYMENT_METHODS = new Set(["cash", "duitnow_qr", "credit_card", "debit_card", "bank_transfer", "e_wallet", "havent_paid"]);
const ORDER_STATUSES = new Set(["paid", "unpaid", "partially_paid", "refunded", "void", "cancelled"]);
const PURCHASE_STATUSES = new Set(["draft", "ordered", "partially_received", "received", "cancelled"]);
const USER_ROLES = new Set(["system_owner", "super_admin", "admin", "manager", "cashier", "inventory_staff", "staff"]);
const USER_STATUSES = new Set(["active", "inactive", "suspended"]);
const DEFAULT_SYSTEM_USERS = [
  {
    username: "TSSC",
    password: "Tssc@123",
    full_name: "System Owner",
    role: "system_owner"
  },
  {
    username: "superadmin",
    password: "Super@123",
    full_name: "Super Administrator",
    role: "super_admin"
  },
  {
    username: "admin",
    password: "Admin@123",
    full_name: "Administrator",
    role: "admin"
  }
];

const DEFAULT_NATURE_OF_BUSINESS = [
  "Convenience Stores",
  "Retail of Football, Hockey, Cricket, Baseball, Badminton, Futsal and Paintball Equipment",
  "Wholesale of a Variety of Goods Without Any Particular Specialization (N.E.C.)"
];

function parseNatureOfBusiness(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item || "").trim()).filter(Boolean);
  }
  const text = String(value || "").trim();
  if (!text) {
    return [...DEFAULT_NATURE_OF_BUSINESS];
  }
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) {
      return parsed.map((item) => String(item || "").trim()).filter(Boolean);
    }
  } catch (_) {
    // Fall back to line-separated/plain text below.
  }
  return text.split(/\r?\n|;/).map((item) => item.trim()).filter(Boolean);
}

function defaultCompanySettings(timestamp = nowISO()) {
  return {
    id: 1,
    company_logo: "/assets/tunas-logo.jpg",
    company_name: "TUNAS BADMINTON SPORTS CENTRE SDN. BHD.",
    business_registration_number: "1536028-U (202301042111)",
    registration_date: "24 October 2023",
    state: "Negeri Sembilan",
    nature_of_business_json: JSON.stringify(DEFAULT_NATURE_OF_BUSINESS),
    business_address: "Lot PT 3962 & 3963, Jln Haruan 2, Pusat Komersial Oakland, 70300 Seremban, Negeri Sembilan",
    phone: "0182063324",
    email: "",
    website: "",
    sst_registration_number: "",
    created_at: timestamp,
    updated_at: timestamp
  };
}

function normalizeCompanySettings(profile = {}, current = defaultCompanySettings()) {
  const timestamp = nowISO();
  const natureOfBusiness = parseNatureOfBusiness(
    Object.hasOwn(profile, "nature_of_business")
      ? profile.nature_of_business
      : profile.nature_of_business_json ?? current.nature_of_business_json
  );
  return {
    id: 1,
    company_logo: String(profile.company_logo ?? current.company_logo ?? "").trim(),
    company_name: String(profile.company_name ?? current.company_name ?? "").trim() || defaultCompanySettings().company_name,
    business_registration_number: String(profile.business_registration_number ?? current.business_registration_number ?? "").trim() || defaultCompanySettings().business_registration_number,
    registration_date: String(profile.registration_date ?? current.registration_date ?? "").trim() || defaultCompanySettings().registration_date,
    state: String(profile.state ?? current.state ?? "").trim() || defaultCompanySettings().state,
    nature_of_business_json: JSON.stringify(natureOfBusiness.length ? natureOfBusiness : DEFAULT_NATURE_OF_BUSINESS),
    business_address: String(profile.business_address ?? current.business_address ?? "").trim(),
    phone: String(profile.phone ?? current.phone ?? "").trim(),
    email: String(profile.email ?? current.email ?? "").trim(),
    website: String(profile.website ?? current.website ?? "").trim(),
    sst_registration_number: String(profile.sst_registration_number ?? current.sst_registration_number ?? "").trim(),
    created_at: String(current.created_at || profile.created_at || timestamp),
    updated_at: timestamp
  };
}

function serializeCompanySettings(profile) {
  return {
    ...profile,
    nature_of_business: parseNatureOfBusiness(profile.nature_of_business_json)
  };
}

function normalizePaymentMethod(value, fallback = "cash") {
  const method = String(value || "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (method === "qr") {
    return "duitnow_qr";
  }
  if (method === "card") {
    return "credit_card";
  }
  if (method === "belum_bayar" || method === "haven't_paid" || method === "haventpay") {
    return "havent_paid";
  }
  return PAYMENT_METHODS.has(method) ? method : fallback;
}

function normalizeOrderStatus(value, fallback = "paid") {
  const status = String(value || "").trim().toLowerCase().replace(/\s+/g, "_");
  if (status === "draft") {
    return "unpaid";
  }
  if (status === "partial" || status === "partially-paid") {
    return "partially_paid";
  }
  if (status === "voided") {
    return "void";
  }
  return ORDER_STATUSES.has(status) ? status : fallback;
}

function normalizePurchaseStatus(value, fallback = "ordered") {
  const status = String(value || "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (status === "partial") {
    return "partially_received";
  }
  return PURCHASE_STATUSES.has(status) ? status : fallback;
}

function hashPassword(value) {
  return crypto.createHash("sha256").update(String(value || ""), "utf8").digest("hex");
}

function hashPasswordSecure(value) {
  const salt = crypto.randomBytes(16).toString("hex");
  const iterations = 120000;
  const hash = crypto.pbkdf2Sync(String(value || ""), salt, iterations, 32, "sha256").toString("hex");
  return `pbkdf2_sha256$${iterations}$${salt}$${hash}`;
}

function isPasswordHashValid(storedHash, plainPassword) {
  const input = String(plainPassword || "");
  const stored = String(storedHash || "");
  if (!input || !stored) {
    return false;
  }

  if (stored.toLowerCase().startsWith("pbkdf2_sha256$")) {
    const [, iterationsText, salt, expected] = stored.split("$");
    const iterations = toNonNegativeInt(iterationsText, 0);
    if (!iterations || !salt || !expected) {
      return false;
    }
    const actual = crypto.pbkdf2Sync(input, salt, iterations, 32, "sha256").toString("hex");
    if (actual.length !== expected.length) {
      return false;
    }
    return crypto.timingSafeEqual(Buffer.from(actual, "hex"), Buffer.from(expected, "hex"));
  }

  if (stored.toLowerCase().startsWith("sha256:")) {
    return stored.toLowerCase() === `sha256:${hashPassword(input)}`;
  }

  return stored === input;
}

function validatePasswordPolicy(password) {
  const value = String(password || "");
  if (value.length < 8 || !/[A-Z]/.test(value) || !/[a-z]/.test(value) || !/[0-9]/.test(value)) {
    const err = new Error("Password must be at least 8 characters and include uppercase, lowercase, and a number.");
    err.code = "BAD_REQUEST";
    throw err;
  }
}

function normalizeUserRole(value, fallback = "cashier") {
  const role = String(value || fallback).trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (role === "cashier_staff") {
    return "cashier";
  }
  return USER_ROLES.has(role) ? role : fallback;
}

function normalizeUserStatus(value, isActive = 1) {
  const status = String(value || "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  if (USER_STATUSES.has(status)) {
    return status;
  }
  return toInt(isActive, 1) ? "active" : "inactive";
}

function isSystemOwnerRole(role) {
  return normalizeUserRole(role, "staff") === "system_owner";
}

function publicUser(user) {
  return {
    id: user.id,
    username: user.username,
    full_name: user.full_name,
    role: user.role,
    email: user.email || "",
    phone: user.phone || "",
    employee_id: user.employee_id || "",
    profile_photo: user.profile_photo || "",
    status: normalizeUserStatus(user.status, user.is_active),
    is_active: normalizeUserStatus(user.status, user.is_active) === "active" ? 1 : 0,
    force_password_change: toInt(user.force_password_change, 0) ? 1 : 0,
    last_login_at: user.last_login_at || "",
    created_at: user.created_at,
    updated_at: user.updated_at || user.created_at
  };
}

function deviceTypeFromUserAgent(userAgent) {
  const value = String(userAgent || "").toLowerCase();
  if (/ipad|tablet/.test(value)) {
    return "Tablet";
  }
  if (/android|iphone|mobile/.test(value)) {
    return "Mobile";
  }
  if (/electron/.test(value)) {
    return "Electron Desktop";
  }
  return "Desktop";
}

function browserFromUserAgent(userAgent) {
  const value = String(userAgent || "");
  if (/Edg\//.test(value)) {
    return "Microsoft Edge";
  }
  if (/Chrome\//.test(value) && !/Edg\//.test(value)) {
    return "Chrome";
  }
  if (/Firefox\//.test(value)) {
    return "Firefox";
  }
  if (/Safari\//.test(value) && !/Chrome\//.test(value)) {
    return "Safari";
  }
  if (/Electron\//.test(value)) {
    return "Electron";
  }
  return value ? "Unknown Browser" : "";
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

function listToValue(value) {
  return listFromValue(value).join(",");
}

function idListFromValue(value) {
  return listFromValue(value)
    .map((item) => toNonNegativeInt(item, 0))
    .filter((item) => item > 0);
}

function normalizeDateString(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : text;
}

class ExcelStore {
  constructor(filePath) {
    this.filePath = filePath;
    this.products = [];
    this.orders = [];
    this.orderItems = [];
    this.payments = [];
    this.pendingOrders = [];
    this.editLogs = [];
    this.users = [];
    this.loginHistory = [];
    this.members = [];
    this.suppliers = [];
    this.purchaseOrders = [];
    this.purchaseOrderItems = [];
    this.expenses = [];
    this.dailyClosings = [];
    this.supplierPayments = [];
    this.discounts = [];
    this.discountUsage = [];
    this.inventoryRecords = [];
    this.outstandingInvoices = [];
    this.customerLedger = [];
    this.companySettings = [];
    this.settingsRows = [];
    this.meta = {
      product_id: 0,
      order_id: 0,
      order_item_id: 0,
      payment_id: 0,
      pending_order_id: 0,
      edit_log_id: 0,
      member_id: 0,
      supplier_id: 0,
      purchase_order_id: 0,
      purchase_order_item_id: 0,
      expense_id: 0,
      daily_closing_id: 0,
      supplier_payment_id: 0,
      discount_id: 0,
      discount_usage_id: 0,
      inventory_record_id: 0,
      outstanding_invoice_id: 0,
      customer_ledger_id: 0,
      login_history_id: 0
    };
  }

  initialize() {
    fs.mkdirSync(path.dirname(this.filePath), { recursive: true });

    if (!fs.existsSync(this.filePath)) {
      this._seedDefaults();
      this._ensureCompanySettingsSeeded();
      this._ensureSettingsSeeded();
      this._save();
      return;
    }

    const workbook = XLSX.readFile(this.filePath, { cellDates: false });
    this.products = this._readSheet(workbook, "products", HEADERS.products).map((row) => ({
      id: toNonNegativeInt(row.id),
      sku: String(row.sku || "").trim(),
      name: String(row.name || "").trim(),
      category: String(row.category || "").trim() || "Uncategorized",
      subcategory: String(row.subcategory || "").trim() || "General",
      barcode: row.barcode ? String(row.barcode).trim() : "",
      brand: String(row.brand || "").trim(),
      supplier: String(row.supplier || "").trim(),
      description: String(row.description || "").trim(),
      template_type: String(row.template_type || "").trim() || "custom",
      attributes_json: String(row.attributes_json || "{}"),
      unit_price_cents: toNonNegativeInt(row.unit_price_cents),
      cost_cents: toNonNegativeInt(row.cost_cents),
      quantity_on_hand: toNonNegativeInt(row.quantity_on_hand),
      reorder_level: toNonNegativeInt(row.reorder_level),
      is_active: toInt(row.is_active, 1) ? 1 : 0,
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.orders = this._readSheet(workbook, "orders", HEADERS.orders).map((row) => ({
      id: toNonNegativeInt(row.id),
      order_number: String(row.order_number || ""),
      invoice_number: String(row.invoice_number || "").trim(),
      invoice_generated_at: String(row.invoice_generated_at || ""),
      invoice_pdf_file_name: String(row.invoice_pdf_file_name || ""),
      invoice_pdf_file_path: String(row.invoice_pdf_file_path || ""),
      invoice_pdf_generated_at: String(row.invoice_pdf_generated_at || ""),
      invoice_print_count: toNonNegativeInt(row.invoice_print_count, 0),
      invoice_export_count: toNonNegativeInt(row.invoice_export_count, 0),
      invoice_last_printed_at: String(row.invoice_last_printed_at || ""),
      invoice_last_exported_at: String(row.invoice_last_exported_at || ""),
      user_id: toNonNegativeInt(row.user_id, 2),
      member_id: toNonNegativeInt(row.member_id, 0),
      customer_name: String(row.customer_name || "").trim(),
      phone: String(row.phone || "").trim(),
      status: normalizeOrderStatus(row.status || "paid"),
      subtotal_cents: toNonNegativeInt(row.subtotal_cents),
      tax_cents: toNonNegativeInt(row.tax_cents),
      discount_cents: toNonNegativeInt(row.discount_cents),
      total_cents: toNonNegativeInt(row.total_cents),
      rounding_cents: toInt(row.rounding_cents, 0),
      payable_cents: toNonNegativeInt(row.payable_cents, toNonNegativeInt(row.total_cents)),
      discount_label: String(row.discount_label || ""),
      notes: String(row.notes || ""),
      created_at: String(row.created_at || nowISO())
    }));

    this.orderItems = this._readSheet(workbook, "order_items", HEADERS.order_items).map((row) => {
      const productId = toNonNegativeInt(row.product_id);
      const product = this.products.find((entry) => Number(entry.id) === Number(productId));
      const fallbackCostCents = toNonNegativeInt(product?.cost_cents, 0);
      return {
        id: toNonNegativeInt(row.id),
        order_id: toNonNegativeInt(row.order_id),
        product_id: productId,
        product_name: String(row.product_name || ""),
        quantity: toNonNegativeInt(row.quantity),
        unit_price_cents: toNonNegativeInt(row.unit_price_cents),
        unit_cost_cents: toNonNegativeInt(row.unit_cost_cents, fallbackCostCents),
        line_total_cents: toNonNegativeInt(row.line_total_cents)
      };
    });

    this.payments = this._readSheet(workbook, "payments", HEADERS.payments).map((row) => ({
      id: toNonNegativeInt(row.id),
      order_id: toNonNegativeInt(row.order_id),
      payment_method: normalizePaymentMethod(row.payment_method || "cash"),
      amount_cents: toNonNegativeInt(row.amount_cents),
      notes: String(row.notes || "").trim(),
      paid_at: String(row.paid_at || nowISO())
    }));

    this.pendingOrders = this._readSheet(workbook, "pending_orders", HEADERS.pending_orders).map((row) => ({
      id: toNonNegativeInt(row.id),
      label: String(row.label || "").trim(),
      customer_name: String(row.customer_name || "").trim(),
      phone: String(row.phone || "").trim(),
      member_id: toNonNegativeInt(row.member_id, 0),
      notes: String(row.notes || "").trim(),
      discount_id: toNonNegativeInt(row.discount_id, 0),
      items_json: String(row.items_json || "[]"),
      status: String(row.status || "active").trim().toLowerCase() === "cancelled" ? "cancelled" : "active",
      created_by: String(row.created_by || "system"),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.editLogs = this._readSheet(workbook, "edit_logs", HEADERS.edit_logs).map((row) => ({
      id: toNonNegativeInt(row.id),
      product_id: toNonNegativeInt(row.product_id),
      action: String(row.action || "edit"),
      changed_by: String(row.changed_by || "system"),
      details: String(row.details || ""),
      created_at: String(row.created_at || nowISO())
    }));

    this.users = this._readSheet(workbook, "users", HEADERS.users).map((row) => ({
      id: toNonNegativeInt(row.id),
      username: String(row.username || "").trim(),
      password: String(row.password || ""),
      full_name: String(row.full_name || "").trim(),
      role: normalizeUserRole(row.role || "staff", "staff"),
      email: String(row.email || "").trim(),
      phone: String(row.phone || "").trim(),
      employee_id: String(row.employee_id || "").trim(),
      profile_photo: String(row.profile_photo || "").trim(),
      status: normalizeUserStatus(row.status, row.is_active),
      is_active: toInt(row.is_active, 1) ? 1 : 0,
      force_password_change: toInt(row.force_password_change, 0) ? 1 : 0,
      last_login_at: String(row.last_login_at || ""),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || row.created_at || nowISO())
    }));

    this.loginHistory = this._readSheet(workbook, "login_history", HEADERS.login_history).map((row) => ({
      id: toNonNegativeInt(row.id),
      user_id: toNonNegativeInt(row.user_id, 0),
      username: String(row.username || "").trim(),
      login_time: String(row.login_time || nowISO()),
      logout_time: String(row.logout_time || ""),
      ip_address: String(row.ip_address || "").trim(),
      device_type: String(row.device_type || "").trim(),
      browser: String(row.browser || "").trim(),
      login_status: String(row.login_status || "success").trim().toLowerCase(),
      session_token_hash: String(row.session_token_hash || "").trim()
    }));

    for (const user of this.users) {
      const stored = String(user.password || "");
      if (!stored) {
        continue;
      }
      const normalizedStored = stored.toLowerCase();
      if (!normalizedStored.startsWith("sha256:") && !normalizedStored.startsWith("pbkdf2_sha256$")) {
        user.password = hashPasswordSecure(stored);
      }
    }

    this.members = this._readSheet(workbook, "members", HEADERS.members).map((row) => ({
      id: toNonNegativeInt(row.id),
      member_no: String(row.member_no || "").trim(),
      name: String(row.name || "").trim(),
      phone: String(row.phone || "").trim(),
      email: String(row.email || "").trim(),
      tier: String(row.tier || "Silver").trim(),
      discount_percent: toNonNegativeInt(row.discount_percent, 0),
      is_active: toInt(row.is_active, 1) ? 1 : 0,
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.suppliers = this._readSheet(workbook, "suppliers", HEADERS.suppliers).map((row) => ({
      id: toNonNegativeInt(row.id),
      name: String(row.name || "").trim(),
      contact_person: String(row.contact_person || "").trim(),
      phone: String(row.phone || "").trim(),
      email: String(row.email || "").trim(),
      address: String(row.address || "").trim(),
      payment_terms: String(row.payment_terms || "").trim(),
      notes: String(row.notes || "").trim(),
      outstanding_cents: toNonNegativeInt(row.outstanding_cents, 0),
      is_active: toInt(row.is_active, 1) ? 1 : 0,
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.purchaseOrders = this._readSheet(workbook, "purchase_orders", HEADERS.purchase_orders).map((row) => ({
      id: toNonNegativeInt(row.id),
      po_number: String(row.po_number || "").trim(),
      supplier_id: toNonNegativeInt(row.supplier_id, 0),
      supplier_name: String(row.supplier_name || "").trim(),
      status: normalizePurchaseStatus(row.status || "ordered"),
      order_date: normalizeDateString(row.order_date || dateKeyLocal()),
      expected_date: normalizeDateString(row.expected_date || ""),
      received_at: String(row.received_at || ""),
      subtotal_cents: toNonNegativeInt(row.subtotal_cents, 0),
      discount_cents: toNonNegativeInt(row.discount_cents, 0),
      total_cents: toNonNegativeInt(row.total_cents, 0),
      notes: String(row.notes || "").trim(),
      created_by: String(row.created_by || "system"),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.purchaseOrderItems = this._readSheet(workbook, "purchase_order_items", HEADERS.purchase_order_items).map((row) => ({
      id: toNonNegativeInt(row.id),
      purchase_order_id: toNonNegativeInt(row.purchase_order_id, 0),
      product_id: toNonNegativeInt(row.product_id, 0),
      product_name: String(row.product_name || "").trim(),
      sku: String(row.sku || "").trim(),
      quantity_ordered: toNonNegativeInt(row.quantity_ordered, 0),
      quantity_received: toNonNegativeInt(row.quantity_received, 0),
      unit_cost_cents: toNonNegativeInt(row.unit_cost_cents, 0),
      line_total_cents: toNonNegativeInt(row.line_total_cents, 0)
    }));

    this.expenses = this._readSheet(workbook, "expenses", HEADERS.expenses).map((row) => ({
      id: toNonNegativeInt(row.id),
      expense_date: normalizeDateString(row.expense_date || dateKeyLocal()),
      category: String(row.category || "General").trim() || "General",
      vendor: String(row.vendor || "").trim(),
      description: String(row.description || "").trim(),
      payment_method: normalizePaymentMethod(row.payment_method || "cash", "cash"),
      amount_cents: toNonNegativeInt(row.amount_cents, 0),
      created_by: String(row.created_by || "system"),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.dailyClosings = this._readSheet(workbook, "daily_closings", HEADERS.daily_closings).map((row) => ({
      id: toNonNegativeInt(row.id),
      closing_date: normalizeDateString(row.closing_date || dateKeyLocal()),
      opening_cash_cents: toNonNegativeInt(row.opening_cash_cents, 0),
      cash_sales_cents: toNonNegativeInt(row.cash_sales_cents, 0),
      cash_expenses_cents: toNonNegativeInt(row.cash_expenses_cents, 0),
      expected_cash_cents: toNonNegativeInt(row.expected_cash_cents, 0),
      counted_cash_cents: toNonNegativeInt(row.counted_cash_cents, 0),
      variance_cents: toInt(row.variance_cents, 0),
      notes: String(row.notes || "").trim(),
      created_by: String(row.created_by || "system"),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.supplierPayments = this._readSheet(workbook, "supplier_payments", HEADERS.supplier_payments).map((row) => ({
      id: toNonNegativeInt(row.id),
      supplier_id: toNonNegativeInt(row.supplier_id, 0),
      supplier_name: String(row.supplier_name || "").trim(),
      payment_method: normalizePaymentMethod(row.payment_method || "cash", "cash"),
      amount_cents: toNonNegativeInt(row.amount_cents, 0),
      notes: String(row.notes || "").trim(),
      paid_at: String(row.paid_at || nowISO()),
      created_by: String(row.created_by || "system")
    }));

    this.discounts = this._readSheet(workbook, "discounts", HEADERS.discounts).map((row) => ({
      id: toNonNegativeInt(row.id),
      name: String(row.name || "").trim(),
      code: String(row.code || "").trim(),
      type: String(row.type || "percentage").trim().toLowerCase(),
      scope: String(row.scope || "cart").trim().toLowerCase(),
      value: Number(row.value || 0),
      start_date: normalizeDateString(row.start_date),
      end_date: normalizeDateString(row.end_date),
      min_purchase_cents: toNonNegativeInt(row.min_purchase_cents, 0),
      max_discount_cents: toNonNegativeInt(row.max_discount_cents, 0),
      applicable_categories: listToValue(row.applicable_categories),
      applicable_product_ids: listToValue(row.applicable_product_ids),
      applicable_member_tiers: listToValue(row.applicable_member_tiers || row.member_tier),
      member_tier: String(row.member_tier || "").trim(),
      promo_type: String(row.promo_type || "").trim(),
      promo_buy_qty: toNonNegativeInt(row.promo_buy_qty, 0),
      promo_free_qty: toNonNegativeInt(row.promo_free_qty, 0),
      promo_discount_percent: Number(row.promo_discount_percent || 0),
      promo_spend_cents: toNonNegativeInt(row.promo_spend_cents, 0),
      promo_discount_cents: toNonNegativeInt(row.promo_discount_cents, 0),
      is_active: toInt(row.is_active, 1) ? 1 : 0,
      created_by: String(row.created_by || "system"),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.discountUsage = this._readSheet(workbook, "discount_usage", HEADERS.discount_usage).map((row) => ({
      id: toNonNegativeInt(row.id),
      discount_id: toNonNegativeInt(row.discount_id, 0),
      discount_name: String(row.discount_name || "").trim(),
      discount_code: String(row.discount_code || "").trim(),
      order_id: toNonNegativeInt(row.order_id, 0),
      order_number: String(row.order_number || "").trim(),
      user_id: toNonNegativeInt(row.user_id, 0),
      staff_name: String(row.staff_name || "").trim(),
      member_id: toNonNegativeInt(row.member_id, 0),
      subtotal_cents: toNonNegativeInt(row.subtotal_cents, 0),
      discount_cents: toNonNegativeInt(row.discount_cents, 0),
      created_at: String(row.created_at || nowISO())
    }));

    this.inventoryRecords = this._readSheet(workbook, "inventory_records", HEADERS.inventory_records).map((row) => ({
      id: toNonNegativeInt(row.id),
      product_id: toNonNegativeInt(row.product_id),
      action: String(row.action || "adjust"),
      quantity_delta: toInt(row.quantity_delta, 0),
      before_qty: toNonNegativeInt(row.before_qty, 0),
      after_qty: toNonNegativeInt(row.after_qty, 0),
      reason: String(row.reason || ""),
      changed_by: String(row.changed_by || "system"),
      created_at: String(row.created_at || nowISO())
    }));

    this.outstandingInvoices = this._readSheet(workbook, "outstanding_invoices", HEADERS.outstanding_invoices).map((row) => ({
      id: toNonNegativeInt(row.id),
      order_id: toNonNegativeInt(row.order_id, 0),
      receipt_number: String(row.receipt_number || "").trim(),
      customer_key: String(row.customer_key || "").trim(),
      customer_name: String(row.customer_name || "").trim() || "Walk-in",
      phone: String(row.phone || "").trim(),
      member_id: toNonNegativeInt(row.member_id, 0),
      description: String(row.description || "").trim() || "Outstanding Invoice",
      original_cents: toNonNegativeInt(row.original_cents, 0),
      paid_cents: toNonNegativeInt(row.paid_cents, 0),
      balance_cents: toNonNegativeInt(row.balance_cents, 0),
      status: normalizeOrderStatus(row.status || "unpaid", "unpaid"),
      created_at: String(row.created_at || nowISO()),
      updated_at: String(row.updated_at || nowISO())
    }));

    this.customerLedger = this._readSheet(workbook, "customer_ledger", HEADERS.customer_ledger).map((row) => ({
      id: toNonNegativeInt(row.id),
      customer_key: String(row.customer_key || "").trim(),
      customer_name: String(row.customer_name || "").trim() || "Walk-in",
      phone: String(row.phone || "").trim(),
      member_id: toNonNegativeInt(row.member_id, 0),
      order_id: toNonNegativeInt(row.order_id, 0),
      receipt_number: String(row.receipt_number || "").trim(),
      description: String(row.description || "").trim(),
      debit_cents: toNonNegativeInt(row.debit_cents, 0),
      credit_cents: toNonNegativeInt(row.credit_cents, 0),
      created_at: String(row.created_at || nowISO()),
      created_by: String(row.created_by || "system")
    }));

    this.companySettings = this._readSheet(workbook, "company_settings", HEADERS.company_settings)
      .map((row) => normalizeCompanySettings(row))
      .filter((row) => row.id);

    this.settingsRows = this._readSheet(workbook, "settings", HEADERS.settings).map((row) => ({
      key: String(row.key || "").trim(),
      value: String(row.value || ""),
      updated_at: String(row.updated_at || nowISO())
    })).filter((row) => row.key);

    const metaRows = this._readSheet(workbook, "meta", HEADERS.meta);
    if (metaRows.length > 0) {
      for (const row of metaRows) {
        const key = String(row.key || "").trim();
        if (Object.hasOwn(this.meta, key)) {
          this.meta[key] = toNonNegativeInt(row.value, 0);
        }
      }
    }

    this._ensureUsersSeeded();
    this._ensureDiscountsSeeded();
    this._ensureSuppliersSeeded();
    this._ensureOutstandingLedgerSeeded();
    this._ensureCompanySettingsSeeded();
    this._ensureSettingsSeeded();
    this._rebuildMeta();
    this._save();
  }

  _getProductRows({
    search = "",
    includeInactive = false,
    category = "",
    subcategory = "",
    brand = "",
    supplier = "",
    status = "",
    stock = "",
    sortBy = "name",
    sortDir = "asc"
  } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    const normalizedCategory = String(category || "").trim().toLowerCase();
    const normalizedSubcategory = String(subcategory || "").trim().toLowerCase();
    const normalizedBrand = String(brand || "").trim().toLowerCase();
    const normalizedSupplier = String(supplier || "").trim().toLowerCase();
    const normalizedStatus = String(status || "").trim().toLowerCase();
    const normalizedStock = String(stock || "").trim().toLowerCase();
    const sortKey = String(sortBy || "name").trim().toLowerCase();
    const direction = String(sortDir || "asc").trim().toLowerCase() === "desc" ? -1 : 1;
    const sortGetters = {
      id: (product) => Number(product.id || 0),
      name: (product) => String(product.name || "").toLowerCase(),
      category: (product) => String(product.category || "").toLowerCase(),
      sku: (product) => String(product.sku || "").toLowerCase(),
      brand: (product) => String(product.brand || "").toLowerCase(),
      cost: (product) => Number(product.cost_cents || 0),
      price: (product) => Number(product.unit_price_cents || 0),
      stock: (product) => Number(product.quantity_on_hand || 0),
      reorder: (product) => Number(product.reorder_level || 0),
      supplier: (product) => String(product.supplier || "").toLowerCase(),
      status: (product) => Number(product.is_active || 0)
    };
    const getter = sortGetters[sortKey] || sortGetters.name;

    return this.products
      .filter((product) => includeInactive || product.is_active)
      .filter((product) => {
        if (!normalizedStatus || normalizedStatus === "all") {
          return true;
        }
        if (normalizedStatus === "active") {
          return Boolean(product.is_active);
        }
        if (normalizedStatus === "inactive") {
          return !product.is_active;
        }
        return true;
      })
      .filter((product) => {
        if (!normalizedStock || normalizedStock === "all") {
          return true;
        }

        const stockOnHand = toNonNegativeInt(product.quantity_on_hand, 0);
        const reorderLevel = toNonNegativeInt(product.reorder_level, 0);
        if (normalizedStock === "out") {
          return stockOnHand <= 0;
        }
        if (normalizedStock === "low") {
          return stockOnHand > 0 && stockOnHand <= reorderLevel;
        }
        if (normalizedStock === "in") {
          return stockOnHand > reorderLevel;
        }
        return true;
      })
      .filter((product) => {
        if (!normalizedCategory) {
          return true;
        }
        return String(product.category || "").trim().toLowerCase() === normalizedCategory;
      })
      .filter((product) => {
        if (!normalizedSubcategory) {
          return true;
        }
        return String(product.subcategory || "").trim().toLowerCase() === normalizedSubcategory;
      })
      .filter((product) => {
        if (!normalizedBrand) {
          return true;
        }
        return String(product.brand || "").trim().toLowerCase() === normalizedBrand;
      })
      .filter((product) => {
        if (!normalizedSupplier) {
          return true;
        }
        return String(product.supplier || "").trim().toLowerCase() === normalizedSupplier;
      })
      .filter((product) => {
        if (!normalized) {
          return true;
        }

        return [
          product.id,
          product.name,
          product.sku,
          product.barcode,
          product.category,
          product.subcategory,
          product.brand,
          product.supplier,
          product.description,
          product.attributes_json
        ]
          .some((field) => String(field || "").toLowerCase().includes(normalized));
      })
      .sort((a, b) => {
        const left = getter(a);
        const right = getter(b);
        if (typeof left === "number" && typeof right === "number") {
          return (left - right || String(a.name || "").localeCompare(String(b.name || ""))) * direction;
        }
        return (String(left).localeCompare(String(right)) || Number(a.id || 0) - Number(b.id || 0)) * direction;
      })
      .map((product) => ({ ...product }));
  }

  getProducts(options = {}) {
    return this._getProductRows(options);
  }

  getProductsPage({
    page = 1,
    pageSize = 10,
    search = "",
    includeInactive = true,
    category = "",
    subcategory = "",
    brand = "",
    supplier = "",
    status = "",
    stock = "",
    sortBy = "name",
    sortDir = "asc"
  } = {}) {
    const safePageSize = Math.min(200, Math.max(1, toNonNegativeInt(pageSize, 10) || 10));
    const rows = this._getProductRows({
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
    const total = rows.length;
    const totalPages = Math.max(1, Math.ceil(total / safePageSize));
    const requestedPage = Math.max(1, toNonNegativeInt(page, 1) || 1);
    const safePage = Math.min(requestedPage, totalPages);
    const start = (safePage - 1) * safePageSize;

    return {
      rows: rows.slice(start, start + safePageSize),
      pagination: {
        page: safePage,
        page_size: safePageSize,
        total,
        total_pages: totalPages
      }
    };
  }

  getProductFilters() {
    const categories = Array.from(
      new Set(this.products.map((product) => String(product.category || "").trim()).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b));

    const brands = Array.from(
      new Set(this.products.map((product) => String(product.brand || "").trim()).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b));

    const suppliers = Array.from(
      new Set(this.products.map((product) => String(product.supplier || "").trim()).filter(Boolean))
    ).sort((a, b) => a.localeCompare(b));

    return { categories, brands, suppliers };
  }

  getProductById(productId) {
    const product = this.products.find((item) => item.id === Number(productId));
    return product ? { ...product } : null;
  }

  getUserByUsername(username, { includeInactive = false } = {}) {
    const normalized = String(username || "").trim().toLowerCase();
    const user = this.users.find((item) => {
      const status = normalizeUserStatus(item.status, item.is_active);
      return item.username.toLowerCase() === normalized && (includeInactive || status === "active");
    });
    return user ? { ...user } : null;
  }

  isPasswordValid(user, plainPassword) {
    return isPasswordHashValid(user?.password, plainPassword);
  }

  getUserById(id, { includeInactive = false } = {}) {
    const user = this.users.find((item) => {
      const status = normalizeUserStatus(item.status, item.is_active);
      return item.id === Number(id) && (includeInactive || status === "active");
    });
    return user ? { ...user } : null;
  }

  getUsers({ includeInactive = true } = {}) {
    return this.users
      .filter((user) => includeInactive || normalizeUserStatus(user.status, user.is_active) === "active")
      .map(publicUser);
  }

  _activeSystemOwnerCount(excludeUserId = 0) {
    return this.users.filter((user) =>
      Number(user.id || 0) !== Number(excludeUserId || 0) &&
      isSystemOwnerRole(user.role) &&
      normalizeUserStatus(user.status, user.is_active) === "active"
    ).length;
  }

  _assertCanChangeSystemOwner(currentUser, nextRole, nextStatus) {
    if (!isSystemOwnerRole(currentUser?.role)) {
      return;
    }
    const stillActiveOwner = isSystemOwnerRole(nextRole) && normalizeUserStatus(nextStatus, currentUser?.is_active) === "active";
    if (stillActiveOwner) {
      return;
    }
    if (this._activeSystemOwnerCount(currentUser.id) < 1) {
      const err = new Error("At least one active System Owner account is required.");
      err.code = "FORBIDDEN";
      throw err;
    }
  }

  addUser(payload, actor = "admin") {
    const username = String(payload.username || "").trim();
    const password = String(payload.password || "").trim();
    const confirmPassword = String(payload.confirm_password || payload.confirmPassword || password);
    const fullName = String(payload.full_name || payload.fullName || "").trim();
    const role = normalizeUserRole(payload.role || "cashier");
    const status = normalizeUserStatus(payload.status, payload.is_active === false ? 0 : 1);
    if (!username || !password) {
      const err = new Error("Username and password are required.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    if (password !== confirmPassword) {
      const err = new Error("Password and confirmation password do not match.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    validatePasswordPolicy(password);
    if (this.users.some((user) => user.username.toLowerCase() === username.toLowerCase())) {
      const err = new Error("Username already exists.");
      err.code = "CONFLICT";
      throw err;
    }

    const createdAt = nowISO();
    const user = {
      id: Math.max(...this.users.map((item) => item.id), 0) + 1,
      username,
      password: hashPasswordSecure(password),
      full_name: fullName || username,
      role,
      email: String(payload.email || "").trim(),
      phone: String(payload.phone || "").trim(),
      employee_id: String(payload.employee_id || payload.employeeId || "").trim(),
      profile_photo: String(payload.profile_photo || payload.profilePhoto || "").trim(),
      status,
      is_active: status === "active" ? 1 : 0,
      force_password_change: toInt(payload.force_password_change, 0) ? 1 : 0,
      last_login_at: "",
      created_at: createdAt,
      updated_at: createdAt
    };
    this.users.push(user);
    this._logAudit({
      action: "create_user",
      module: "User Management",
      changedBy: actor,
      description: `Added ${role} user ${username}`,
      details: { after: publicUser(user) }
    });
    this._save();
    return publicUser(user);
  }

  updateUser(userId, payload, actor = "admin") {
    const id = Number(userId);
    const idx = this.users.findIndex((user) => Number(user.id) === id);
    if (idx < 0) {
      const err = new Error("User not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const currentUser = this.users[idx];
    const before = { ...currentUser, password: "" };
    const nextStatus = Object.hasOwn(payload, "status")
      ? normalizeUserStatus(payload.status, currentUser.is_active)
      : Object.hasOwn(payload, "is_active")
        ? normalizeUserStatus("", payload.is_active === false ? 0 : 1)
        : normalizeUserStatus(currentUser.status, currentUser.is_active);
    const nextRole = normalizeUserRole(payload.role || currentUser.role || "cashier");
    this._assertCanChangeSystemOwner(currentUser, nextRole, nextStatus);
    if (Object.hasOwn(payload, "username")) {
      const nextUsername = String(payload.username || "").trim();
      if (!nextUsername) {
        const err = new Error("Username is required.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      if (this.users.some((user) => Number(user.id) !== id && String(user.username || "").toLowerCase() === nextUsername.toLowerCase())) {
        const err = new Error("Username already exists.");
        err.code = "CONFLICT";
        throw err;
      }
      this.users[idx].username = nextUsername;
    }
    this.users[idx].full_name = String(payload.full_name || payload.fullName || currentUser.full_name).trim();
    this.users[idx].role = nextRole;
    this.users[idx].email = String(Object.hasOwn(payload, "email") ? payload.email || "" : currentUser.email || "").trim();
    this.users[idx].phone = String(Object.hasOwn(payload, "phone") ? payload.phone || "" : currentUser.phone || "").trim();
    this.users[idx].employee_id = String(
      Object.hasOwn(payload, "employee_id") ? payload.employee_id || "" : payload.employeeId || currentUser.employee_id || ""
    ).trim();
    this.users[idx].profile_photo = String(
      Object.hasOwn(payload, "profile_photo") ? payload.profile_photo || "" : payload.profilePhoto || currentUser.profile_photo || ""
    ).trim();
    this.users[idx].status = nextStatus;
    this.users[idx].is_active = nextStatus === "active" ? 1 : 0;
    if (Object.hasOwn(payload, "force_password_change")) {
      this.users[idx].force_password_change = toInt(payload.force_password_change, 0) ? 1 : 0;
    }
    this.users[idx].updated_at = nowISO();
    if (payload.password) {
      validatePasswordPolicy(payload.password);
      if (payload.confirm_password || payload.confirmPassword) {
        const confirmPassword = String(payload.confirm_password || payload.confirmPassword || "");
        if (String(payload.password) !== confirmPassword) {
          const err = new Error("Password and confirmation password do not match.");
          err.code = "BAD_REQUEST";
          throw err;
        }
      }
      this.users[idx].password = hashPasswordSecure(payload.password);
    }
    this._logAudit({
      action: "edit_user",
      module: "User Management",
      changedBy: actor,
      description: `Updated user ${this.users[idx].username}`,
      details: { before, after: { ...this.users[idx], password: "" } }
    });
    this._save();
    return publicUser(this.users[idx]);
  }

  resetUserPassword(userId, payload, actor = "admin") {
    const id = Number(userId);
    const idx = this.users.findIndex((user) => Number(user.id) === id);
    if (idx < 0) {
      const err = new Error("User not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const password = String(payload.password || "").trim();
    const confirmPassword = String(payload.confirm_password || payload.confirmPassword || "").trim();
    if (!password || password !== confirmPassword) {
      const err = new Error("Password and confirmation password are required and must match.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    validatePasswordPolicy(password);
    this.users[idx].password = hashPasswordSecure(password);
    this.users[idx].force_password_change = Object.hasOwn(payload, "force_password_change")
      ? toInt(payload.force_password_change, 0) ? 1 : 0
      : 0;
    this.users[idx].updated_at = nowISO();
    this._logAudit({
      action: "reset_password",
      module: "User Management",
      changedBy: actor,
      description: `Reset password for user ${this.users[idx].username}`,
      details: { user_id: id, username: this.users[idx].username }
    });
    this._save();
    return publicUser(this.users[idx]);
  }

  removeUser(userId, actor = "admin") {
    const id = Number(userId);
    const idx = this.users.findIndex((user) => Number(user.id) === id);
    if (idx < 0) {
      const err = new Error("User not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const before = { ...this.users[idx], password: "" };
    if (isSystemOwnerRole(this.users[idx].role)) {
      const err = new Error("System Owner accounts cannot be deleted.");
      err.code = "FORBIDDEN";
      throw err;
    }
    this.users[idx].status = "inactive";
    this.users[idx].is_active = 0;
    this.users[idx].updated_at = nowISO();
    this._logAudit({
      action: "delete_user",
      module: "User Management",
      changedBy: actor,
      description: `Deactivated user ${this.users[idx].username}`,
      details: { before, after: { ...this.users[idx], password: "" } }
    });
    this._save();
    return publicUser(this.users[idx]);
  }

  recordLoginAttempt({ userId = 0, username, tokenHash = "", ipAddress = "", userAgent = "", status = "success" } = {}) {
    const loginTime = nowISO();
    const row = {
      id: this._nextId("login_history_id"),
      user_id: Number(userId || 0),
      username: String(username || "").trim(),
      login_time: loginTime,
      logout_time: "",
      ip_address: String(ipAddress || "").trim(),
      device_type: deviceTypeFromUserAgent(userAgent),
      browser: browserFromUserAgent(userAgent),
      login_status: String(status || "success").trim().toLowerCase(),
      session_token_hash: String(tokenHash || "").trim()
    };
    this.loginHistory.push(row);
    if (row.user_id && row.login_status === "success") {
      const user = this.users.find((item) => Number(item.id) === Number(row.user_id));
      if (user) {
        user.last_login_at = loginTime;
        user.updated_at = user.updated_at || loginTime;
      }
    }
    this._save();
    return { ...row };
  }

  recordLogout(tokenHash, logoutTime = nowISO()) {
    const hash = String(tokenHash || "").trim();
    const row = [...this.loginHistory].reverse().find((item) => String(item.session_token_hash || "") === hash && !item.logout_time);
    if (row) {
      row.logout_time = logoutTime;
      this._save();
    }
    return row ? { ...row } : null;
  }

  getLoginHistory({ limit = 500, username = "" } = {}) {
    const normalized = String(username || "").trim().toLowerCase();
    const max = Math.min(1000, Math.max(1, toNonNegativeInt(limit, 500) || 500));
    return this.loginHistory
      .filter((row) => !normalized || String(row.username || "").toLowerCase().includes(normalized))
      .slice()
      .sort((a, b) => Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((row) => ({ ...row }));
  }

  getSetting(key, fallback = "") {
    const row = this.settingsRows.find((item) => String(item.key || "") === String(key || ""));
    return row ? String(row.value || "") : fallback;
  }

  _setSettingValue(key, value) {
    const settingKey = String(key || "").trim();
    const settingValue = String(value || "");
    const now = nowISO();
    const row = this.settingsRows.find((item) => String(item.key || "") === settingKey);
    if (row) {
      row.value = settingValue;
      row.updated_at = now;
    } else {
      this.settingsRows.push({ key: settingKey, value: settingValue, updated_at: now });
    }
    return { key: settingKey, value: settingValue, updated_at: now };
  }

  setSetting(key, value, actor = "system") {
    const settingKey = String(key || "").trim();
    const result = this._setSettingValue(settingKey, value);
    this._logAudit({
      action: "system_setting_changed",
      module: "Settings",
      changedBy: actor,
      description: `Updated setting ${settingKey}`,
      details: { key: settingKey }
    });
    this._save();
    return result;
  }

  getCompanySettings() {
    this._ensureCompanySettingsSeeded();
    return serializeCompanySettings(this.companySettings[0] || defaultCompanySettings());
  }

  updateCompanySettings(profile = {}, actor = "system") {
    const current = this.companySettings[0] || defaultCompanySettings();
    const next = normalizeCompanySettings(profile, current);
    this.companySettings = [next];
    this._logAudit({
      action: "company_profile_changed",
      module: "Settings",
      changedBy: actor,
      description: "Updated company information",
      details: {
        company_name: next.company_name,
        business_registration_number: next.business_registration_number
      }
    });
    this._save();
    return this.getCompanySettings();
  }

  getMembers({ search = "", includeInactive = false } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    return this.members
      .filter((member) => includeInactive || member.is_active)
      .filter((member) => {
        if (!normalized) {
          return true;
        }
        return [member.member_no, member.name, member.phone, member.email, member.tier].some((field) =>
          String(field || "").toLowerCase().includes(normalized)
        );
      })
      .slice()
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
      .map((member) => ({ ...member }));
  }

  getMemberById(id) {
    const member = this.members.find((item) => Number(item.id) === Number(id) && item.is_active);
    return member ? { ...member } : null;
  }

  addMember(payload, actor = "system") {
    const name = String(payload.name || "").trim();
    const phone = String(payload.phone || "").trim();
    if (!name) {
      const err = new Error("Member name is required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const duplicate = phone
      ? this.members.find((member) => String(member.phone || "").trim() === phone && member.is_active)
      : null;
    if (duplicate) {
      const err = new Error("Member phone already exists.");
      err.code = "CONFLICT";
      throw err;
    }

    const id = this._nextId("member_id");
    const timestamp = nowISO();
    const tier = String(payload.tier || "Silver").trim() || "Silver";
    const member = {
      id,
      member_no: String(payload.member_no || `MBR-${String(id).padStart(5, "0")}`).trim(),
      name,
      phone,
      email: String(payload.email || "").trim(),
      tier,
      discount_percent: toNonNegativeInt(payload.discount_percent, tier === "VIP" ? 15 : tier === "Gold" ? 10 : 5),
      is_active: payload.is_active === false ? 0 : 1,
      created_at: timestamp,
      updated_at: timestamp
    };
    this.members.push(member);
    this._logEdit({
      productId: 0,
      action: "member_add",
      changedBy: actor,
      details: JSON.stringify({ after: member })
    });
    this._save();
    return { ...member };
  }

  updateMember(memberId, payload, actor = "system") {
    const id = Number(memberId);
    const idx = this.members.findIndex((member) => Number(member.id) === id);
    if (idx < 0) {
      const err = new Error("Member not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const current = this.members[idx];
    const name = String(payload.name || "").trim();
    if (!name) {
      const err = new Error("Member name is required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const before = { ...current };
    const tier = String(payload.tier || current.tier || "Silver").trim() || "Silver";
    const updated = {
      ...current,
      member_no: String(payload.member_no || current.member_no || "").trim(),
      name,
      phone: String(payload.phone || "").trim(),
      email: String(payload.email || "").trim(),
      tier,
      discount_percent: toNonNegativeInt(payload.discount_percent, current.discount_percent || 0),
      is_active: payload.is_active === false ? 0 : 1,
      updated_at: nowISO()
    };
    this.members[idx] = updated;
    this._logEdit({
      productId: 0,
      action: "member_edit",
      changedBy: actor,
      details: JSON.stringify({ before, after: updated })
    });
    this._save();
    return { ...updated };
  }

  removeMember(memberId, actor = "system") {
    const id = Number(memberId);
    const idx = this.members.findIndex((member) => Number(member.id) === id);
    if (idx < 0) {
      const err = new Error("Member not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const before = { ...this.members[idx] };
    this.members[idx].is_active = 0;
    this.members[idx].updated_at = nowISO();
    this._logEdit({
      productId: 0,
      action: "member_remove",
      changedBy: actor,
      details: JSON.stringify({ before, after: this.members[idx] })
    });
    this._save();
    return { ...this.members[idx] };
  }

  getSuppliers({ search = "", includeInactive = false } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    return this.suppliers
      .filter((supplier) => includeInactive || supplier.is_active)
      .filter((supplier) => {
        if (!normalized) {
          return true;
        }
        return [
          supplier.name,
          supplier.contact_person,
          supplier.phone,
          supplier.email,
          supplier.payment_terms,
          supplier.notes
        ].some((field) => String(field || "").toLowerCase().includes(normalized));
      })
      .slice()
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
      .map((supplier) => ({ ...supplier }));
  }

  getSupplierById(id, { includeInactive = false } = {}) {
    const supplier = this.suppliers.find((item) => Number(item.id) === Number(id) && (includeInactive || item.is_active));
    return supplier ? { ...supplier } : null;
  }

  _normalizeSupplierPayload(payload = {}, current = null) {
    const name = String(payload.name ?? current?.name ?? "").trim();
    if (!name) {
      const err = new Error("Supplier name is required.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    const duplicate = this.suppliers.find((supplier) =>
      Number(supplier.id) !== Number(current?.id || 0) &&
      supplier.is_active &&
      String(supplier.name || "").trim().toLowerCase() === name.toLowerCase()
    );
    if (duplicate) {
      const err = new Error("Supplier name already exists.");
      err.code = "CONFLICT";
      throw err;
    }
    const timestamp = nowISO();
    return {
      id: current ? current.id : this._nextId("supplier_id"),
      name,
      contact_person: String(payload.contact_person ?? current?.contact_person ?? "").trim(),
      phone: String(payload.phone ?? current?.phone ?? "").trim(),
      email: String(payload.email ?? current?.email ?? "").trim(),
      address: String(payload.address ?? current?.address ?? "").trim(),
      payment_terms: String(payload.payment_terms ?? current?.payment_terms ?? "COD").trim() || "COD",
      notes: String(payload.notes ?? current?.notes ?? "").trim(),
      outstanding_cents: toNonNegativeInt(current?.outstanding_cents ?? payload.outstanding_cents ?? 0, 0),
      is_active: payload.is_active === false ? 0 : 1,
      created_at: current?.created_at || timestamp,
      updated_at: timestamp
    };
  }

  addSupplier(payload, actor = "system") {
    const supplier = this._normalizeSupplierPayload(payload);
    this.suppliers.push(supplier);
    this._logAudit({
      action: "supplier_created",
      module: "Suppliers",
      changedBy: actor,
      description: `Created supplier ${supplier.name}`,
      details: { after: supplier }
    });
    this._save();
    return { ...supplier };
  }

  updateSupplier(supplierId, payload, actor = "system") {
    const id = Number(supplierId);
    const idx = this.suppliers.findIndex((supplier) => Number(supplier.id) === id);
    if (idx < 0) {
      const err = new Error("Supplier not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const before = { ...this.suppliers[idx] };
    const updated = this._normalizeSupplierPayload(payload, before);
    this.suppliers[idx] = updated;
    this._logAudit({
      action: "supplier_updated",
      module: "Suppliers",
      changedBy: actor,
      description: `Updated supplier ${updated.name}`,
      details: { before, after: updated }
    });
    this._save();
    return { ...updated };
  }

  removeSupplier(supplierId, actor = "system") {
    const id = Number(supplierId);
    const idx = this.suppliers.findIndex((supplier) => Number(supplier.id) === id);
    if (idx < 0) {
      const err = new Error("Supplier not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const before = { ...this.suppliers[idx] };
    this.suppliers[idx].is_active = 0;
    this.suppliers[idx].updated_at = nowISO();
    this._logAudit({
      action: "supplier_deactivated",
      module: "Suppliers",
      changedBy: actor,
      description: `Deactivated supplier ${before.name}`,
      details: { before, after: this.suppliers[idx] }
    });
    this._save();
    return { ...this.suppliers[idx] };
  }

  _purchaseItemsForOrder(purchaseOrderId) {
    return this.purchaseOrderItems
      .filter((item) => Number(item.purchase_order_id) === Number(purchaseOrderId))
      .map((item) => ({ ...item }));
  }

  _mapPurchaseOrder(order) {
    const items = this._purchaseItemsForOrder(order.id);
    const supplier = this.getSupplierById(order.supplier_id, { includeInactive: true });
    const orderedQty = items.reduce((sum, item) => sum + toNonNegativeInt(item.quantity_ordered, 0), 0);
    const receivedQty = items.reduce((sum, item) => sum + toNonNegativeInt(item.quantity_received, 0), 0);
    return {
      ...order,
      supplier,
      items,
      item_count: items.length,
      quantity_ordered: orderedQty,
      quantity_received: receivedQty
    };
  }

  getPurchaseOrders({ search = "", status = "", includeCancelled = false, limit = 500 } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    const normalizedStatus = normalizePurchaseStatus(status || "", "");
    const max = Math.min(1000, Math.max(1, toNonNegativeInt(limit, 500) || 500));
    return this.purchaseOrders
      .filter((order) => includeCancelled || order.status !== "cancelled")
      .filter((order) => !normalizedStatus || order.status === normalizedStatus)
      .filter((order) => {
        if (!normalized) {
          return true;
        }
        const items = this._purchaseItemsForOrder(order.id);
        return [
          order.po_number,
          order.supplier_name,
          order.status,
          order.notes,
          ...items.flatMap((item) => [item.product_name, item.sku])
        ].some((field) => String(field || "").toLowerCase().includes(normalized));
      })
      .slice()
      .sort((a, b) => String(b.created_at || "").localeCompare(String(a.created_at || "")) || Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((order) => this._mapPurchaseOrder(order));
  }

  getPurchaseOrderById(id) {
    const order = this.purchaseOrders.find((entry) => Number(entry.id) === Number(id));
    return order ? this._mapPurchaseOrder(order) : null;
  }

  createPurchaseOrder(payload, actor = "system") {
    const supplierId = toNonNegativeInt(payload.supplier_id, 0);
    const supplier = this.suppliers.find((item) => Number(item.id) === supplierId && item.is_active);
    if (!supplier) {
      const err = new Error("Active supplier is required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const rawItems = Array.isArray(payload.items) ? payload.items : [];
    if (!rawItems.length) {
      const err = new Error("At least one purchase item is required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const timestamp = nowISO();
    const orderId = this._nextId("purchase_order_id");
    const items = rawItems.map((raw) => {
      const productId = toNonNegativeInt(raw.product_id, 0);
      const product = this.products.find((entry) => Number(entry.id) === productId && entry.is_active);
      if (!product) {
        const err = new Error("Each purchase item must reference an active product.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      const quantity = toNonNegativeInt(raw.quantity_ordered ?? raw.quantity, 0);
      if (quantity <= 0) {
        const err = new Error("Purchase quantity must be greater than zero.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      const unitCostCents = centsFromValue(raw.unit_cost ?? raw.cost ?? raw.cost_price ?? product.cost_cents / 100);
      if (unitCostCents <= 0) {
        const err = new Error("Unit cost must be greater than zero.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      return {
        id: this._nextId("purchase_order_item_id"),
        purchase_order_id: orderId,
        product_id: product.id,
        product_name: product.name,
        sku: product.sku || "",
        quantity_ordered: quantity,
        quantity_received: 0,
        unit_cost_cents: unitCostCents,
        line_total_cents: quantity * unitCostCents
      };
    });

    const subtotalCents = items.reduce((sum, item) => sum + toNonNegativeInt(item.line_total_cents, 0), 0);
    const discountCents = Math.min(subtotalCents, centsFromValue(payload.discount ?? payload.discount_amount ?? 0));
    const order = {
      id: orderId,
      po_number: String(payload.po_number || purchaseOrderNumber(orderId, timestamp)).trim(),
      supplier_id: supplier.id,
      supplier_name: supplier.name,
      status: normalizePurchaseStatus(payload.status || "ordered"),
      order_date: normalizeDateString(payload.order_date || dateKeyLocal()),
      expected_date: normalizeDateString(payload.expected_date || ""),
      received_at: "",
      subtotal_cents: subtotalCents,
      discount_cents: discountCents,
      total_cents: Math.max(0, subtotalCents - discountCents),
      notes: String(payload.notes || "").trim(),
      created_by: String(actor || "system"),
      created_at: timestamp,
      updated_at: timestamp
    };

    this.purchaseOrders.push(order);
    this.purchaseOrderItems.push(...items);
    this._logAudit({
      action: "purchase_order_created",
      module: "Purchasing",
      changedBy: actor,
      description: `Created purchase order ${order.po_number}`,
      details: { order, items }
    });
    this._save();
    return this._mapPurchaseOrder(order);
  }

  receivePurchaseOrder(purchaseOrderId, payload = {}, actor = "system") {
    const id = Number(purchaseOrderId);
    const orderIdx = this.purchaseOrders.findIndex((order) => Number(order.id) === id);
    if (orderIdx < 0) {
      const err = new Error("Purchase order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const order = this.purchaseOrders[orderIdx];
    if (order.status === "cancelled") {
      const err = new Error("Cancelled purchase orders cannot be received.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const requestedItems = new Map(
      (Array.isArray(payload.items) ? payload.items : []).map((item) => [
        Number(item.id || item.purchase_order_item_id || 0),
        toNonNegativeInt(item.receive_quantity ?? item.quantity_received ?? item.quantity, 0)
      ])
    );
    const receiveAll = requestedItems.size === 0 || payload.receive_all === true;
    const changedItems = [];
    let receivedValueCents = 0;

    for (const item of this.purchaseOrderItems.filter((entry) => Number(entry.purchase_order_id) === id)) {
      const remaining = Math.max(0, toNonNegativeInt(item.quantity_ordered, 0) - toNonNegativeInt(item.quantity_received, 0));
      if (!receiveAll && !requestedItems.has(Number(item.id))) {
        continue;
      }
      const requested = requestedItems.has(Number(item.id)) ? requestedItems.get(Number(item.id)) : remaining;
      const quantity = receiveAll && !requestedItems.has(Number(item.id)) ? remaining : Math.min(remaining, requested);
      if (quantity <= 0) {
        continue;
      }

      const productIdx = this.products.findIndex((product) => Number(product.id) === Number(item.product_id));
      if (productIdx < 0) {
        const err = new Error(`Product not found for ${item.product_name}.`);
        err.code = "NOT_FOUND";
        throw err;
      }

      const beforeQty = toNonNegativeInt(this.products[productIdx].quantity_on_hand, 0);
      this.products[productIdx].quantity_on_hand = beforeQty + quantity;
      if (toNonNegativeInt(item.unit_cost_cents, 0) > 0) {
        this.products[productIdx].cost_cents = toNonNegativeInt(item.unit_cost_cents, 0);
      }
      this.products[productIdx].updated_at = nowISO();
      item.quantity_received = toNonNegativeInt(item.quantity_received, 0) + quantity;
      receivedValueCents += quantity * toNonNegativeInt(item.unit_cost_cents, 0);
      changedItems.push({ ...item, received_now: quantity });
      this._logInventoryRecord({
        productId: item.product_id,
        action: "purchase_receive",
        quantityDelta: quantity,
        beforeQty,
        afterQty: this.products[productIdx].quantity_on_hand,
        reason: `Received ${order.po_number}`,
        changedBy: actor
      });
    }

    if (!changedItems.length) {
      const err = new Error("No remaining purchase quantities to receive.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const allItems = this.purchaseOrderItems.filter((entry) => Number(entry.purchase_order_id) === id);
    const totalOrdered = allItems.reduce((sum, item) => sum + toNonNegativeInt(item.quantity_ordered, 0), 0);
    const totalReceived = allItems.reduce((sum, item) => sum + toNonNegativeInt(item.quantity_received, 0), 0);
    order.status = totalReceived >= totalOrdered ? "received" : "partially_received";
    order.received_at = order.status === "received" ? nowISO() : order.received_at || "";
    order.updated_at = nowISO();

    const supplierIdx = this.suppliers.findIndex((supplier) => Number(supplier.id) === Number(order.supplier_id));
    if (supplierIdx >= 0) {
      this.suppliers[supplierIdx].outstanding_cents = toNonNegativeInt(this.suppliers[supplierIdx].outstanding_cents, 0) + receivedValueCents;
      this.suppliers[supplierIdx].updated_at = nowISO();
    }

    this._logAudit({
      action: "purchase_order_received",
      module: "Purchasing",
      changedBy: actor,
      description: `Received stock for ${order.po_number}`,
      details: { order_id: order.id, received_value_cents: receivedValueCents, items: changedItems }
    });
    this._save();
    return this._mapPurchaseOrder(order);
  }

  getPurchaseHistory({ search = "", limit = 500 } = {}) {
    return this.getPurchaseOrders({ search, includeCancelled: true, limit })
      .filter((order) => ["partially_received", "received"].includes(order.status));
  }

  getExpenses({ start_date = "", end_date = "", search = "", limit = 500 } = {}) {
    const start = normalizeDateString(start_date || "");
    const end = normalizeDateString(end_date || "");
    const normalized = String(search || "").trim().toLowerCase();
    const max = Math.min(1000, Math.max(1, toNonNegativeInt(limit, 500) || 500));
    return this.expenses
      .filter((expense) => !start || String(expense.expense_date || "") >= start)
      .filter((expense) => !end || String(expense.expense_date || "") <= end)
      .filter((expense) => {
        if (!normalized) {
          return true;
        }
        return [expense.category, expense.vendor, expense.description, expense.payment_method, expense.created_by].some((field) =>
          String(field || "").toLowerCase().includes(normalized)
        );
      })
      .slice()
      .sort((a, b) => String(b.expense_date || "").localeCompare(String(a.expense_date || "")) || Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((expense) => ({ ...expense }));
  }

  addExpense(payload, actor = "system") {
    const amountCents = centsFromValue(payload.amount ?? payload.amount_paid ?? 0);
    if (amountCents <= 0) {
      const err = new Error("Expense amount must be greater than zero.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    const timestamp = nowISO();
    const expense = {
      id: this._nextId("expense_id"),
      expense_date: normalizeDateString(payload.expense_date || dateKeyLocal()),
      category: String(payload.category || "General").trim() || "General",
      vendor: String(payload.vendor || "").trim(),
      description: String(payload.description || "").trim(),
      payment_method: normalizePaymentMethod(payload.payment_method || "cash", "cash"),
      amount_cents: amountCents,
      created_by: String(actor || "system"),
      created_at: timestamp,
      updated_at: timestamp
    };
    this.expenses.push(expense);
    this._logAudit({
      action: "expense_recorded",
      module: "Finance",
      changedBy: actor,
      description: `Recorded ${expense.category} expense`,
      details: { expense }
    });
    this._save();
    return { ...expense };
  }

  getSupplierPayables({ search = "" } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    return this.suppliers
      .filter((supplier) => toNonNegativeInt(supplier.outstanding_cents, 0) > 0)
      .filter((supplier) => {
        if (!normalized) {
          return true;
        }
        return [supplier.name, supplier.contact_person, supplier.phone, supplier.email].some((field) =>
          String(field || "").toLowerCase().includes(normalized)
        );
      })
      .slice()
      .sort((a, b) => toNonNegativeInt(b.outstanding_cents, 0) - toNonNegativeInt(a.outstanding_cents, 0))
      .map((supplier) => ({ ...supplier }));
  }

  getSupplierPayments({ limit = 100 } = {}) {
    const max = Math.min(500, Math.max(1, toNonNegativeInt(limit, 100) || 100));
    return this.supplierPayments
      .slice()
      .sort((a, b) => String(b.paid_at || "").localeCompare(String(a.paid_at || "")) || Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((payment) => ({ ...payment }));
  }

  recordSupplierPayment(supplierId, payload, actor = "system") {
    const id = Number(supplierId);
    const idx = this.suppliers.findIndex((supplier) => Number(supplier.id) === id);
    if (idx < 0) {
      const err = new Error("Supplier not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const amountCents = centsFromValue(payload.amount ?? payload.amount_paid ?? 0);
    if (amountCents <= 0) {
      const err = new Error("Payment amount must be greater than zero.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    const outstandingCents = toNonNegativeInt(this.suppliers[idx].outstanding_cents, 0);
    if (amountCents > outstandingCents) {
      const err = new Error("Payment amount cannot exceed outstanding payable.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    const payment = {
      id: this._nextId("supplier_payment_id"),
      supplier_id: this.suppliers[idx].id,
      supplier_name: this.suppliers[idx].name,
      payment_method: normalizePaymentMethod(payload.payment_method || "bank_transfer", "bank_transfer"),
      amount_cents: amountCents,
      notes: String(payload.notes || "").trim(),
      paid_at: nowISO(),
      created_by: String(actor || "system")
    };
    const before = { ...this.suppliers[idx] };
    this.suppliers[idx].outstanding_cents = Math.max(0, outstandingCents - amountCents);
    this.suppliers[idx].updated_at = nowISO();
    this.supplierPayments.push(payment);
    this._logAudit({
      action: "supplier_payment_recorded",
      module: "Finance",
      changedBy: actor,
      description: `Recorded supplier payment for ${payment.supplier_name}`,
      details: { before, after: this.suppliers[idx], payment }
    });
    this._save();
    return { payment: { ...payment }, supplier: { ...this.suppliers[idx] } };
  }

  getDailyClosings({ limit = 100 } = {}) {
    const max = Math.min(500, Math.max(1, toNonNegativeInt(limit, 100) || 100));
    return this.dailyClosings
      .slice()
      .sort((a, b) => String(b.closing_date || "").localeCompare(String(a.closing_date || "")) || Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((closing) => ({ ...closing }));
  }

  saveDailyClosing(payload, actor = "system") {
    const closingDate = normalizeDateString(payload.closing_date || dateKeyLocal());
    const summary = this.getTodaySummary(closingDate);
    const expenses = this.getExpenses({ start_date: closingDate, end_date: closingDate, limit: 1000 });
    const openingCashCents = centsFromValue(payload.opening_cash ?? payload.opening_cash_amount ?? 0);
    const countedCashCents = centsFromValue(payload.counted_cash ?? payload.counted_cash_amount ?? 0);
    const cashSalesCents = toNonNegativeInt(summary.payment_methods_cents?.cash, 0);
    const cashExpensesCents = expenses
      .filter((expense) => normalizePaymentMethod(expense.payment_method, "") === "cash")
      .reduce((sum, expense) => sum + toNonNegativeInt(expense.amount_cents, 0), 0);
    const expectedCashCents = Math.max(0, openingCashCents + cashSalesCents - cashExpensesCents);
    const timestamp = nowISO();
    const existingIdx = this.dailyClosings.findIndex((closing) => String(closing.closing_date || "") === closingDate);
    const closing = {
      id: existingIdx >= 0 ? this.dailyClosings[existingIdx].id : this._nextId("daily_closing_id"),
      closing_date: closingDate,
      opening_cash_cents: openingCashCents,
      cash_sales_cents: cashSalesCents,
      cash_expenses_cents: cashExpensesCents,
      expected_cash_cents: expectedCashCents,
      counted_cash_cents: countedCashCents,
      variance_cents: countedCashCents - expectedCashCents,
      notes: String(payload.notes || "").trim(),
      created_by: String(actor || "system"),
      created_at: existingIdx >= 0 ? this.dailyClosings[existingIdx].created_at : timestamp,
      updated_at: timestamp
    };
    if (existingIdx >= 0) {
      this.dailyClosings[existingIdx] = closing;
    } else {
      this.dailyClosings.push(closing);
    }
    this._logAudit({
      action: "daily_closing_saved",
      module: "Finance",
      changedBy: actor,
      description: `Saved daily closing for ${closingDate}`,
      details: { closing }
    });
    this._save();
    return { ...closing };
  }

  getFinanceSummary({ date = "" } = {}) {
    const targetDate = normalizeDateString(date || dateKeyLocal());
    const sales = this.getTodaySummary(targetDate);
    const outstanding = this.getOutstandingSummary(5);
    const expenses = this.getExpenses({ start_date: targetDate, end_date: targetDate, limit: 1000 });
    const expenseCents = expenses.reduce((sum, expense) => sum + toNonNegativeInt(expense.amount_cents, 0), 0);
    const payablesCents = this.suppliers.reduce((sum, supplier) => sum + toNonNegativeInt(supplier.outstanding_cents, 0), 0);
    const supplierPayments = this.getSupplierPayments({ limit: 5 });
    return {
      date: targetDate,
      revenue_cents: toNonNegativeInt(sales.total_cents, 0),
      cash_sales_cents: toNonNegativeInt(sales.payment_methods_cents?.cash, 0),
      expense_cents: expenseCents,
      income_cents: toNonNegativeInt(sales.total_cents, 0) - expenseCents,
      outstanding_receivable_cents: toNonNegativeInt(outstanding.outstanding_cents, 0),
      outstanding_payable_cents: payablesCents,
      unpaid_sales_cents: toNonNegativeInt(outstanding.todays_unpaid_cents, 0),
      order_count: toNonNegativeInt(sales.order_count, 0),
      recent_supplier_payments: supplierPayments,
      recent_expenses: expenses.slice(0, 5),
      recent_payments: outstanding.recent_payments || []
    };
  }

  _normalizeDiscountPayload(payload = {}, current = null, actor = "admin") {
    const name = String(payload.name ?? current?.name ?? "").trim();
    const type = String(payload.type ?? current?.type ?? "percentage").trim().toLowerCase();
    const value = Number(payload.value ?? current?.value ?? 0);
    const scope = String(payload.scope ?? current?.scope ?? "cart").trim().toLowerCase();
    const applicableMemberTiers = listToValue(
      payload.applicable_member_tiers ?? payload.member_tier ?? current?.applicable_member_tiers ?? current?.member_tier ?? ""
    );

    if (!name) {
      const err = new Error("Discount name is required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    if (!["percentage", "fixed", "fixed_per_unit", "member", "product", "category", "promotion"].includes(type)) {
      const err = new Error("Invalid discount type.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    if (type !== "promotion" && (!Number.isFinite(value) || value <= 0)) {
      const err = new Error("Discount value must be greater than zero.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const timestamp = nowISO();
    return {
      id: current ? current.id : this._nextId("discount_id"),
      name,
      code: String(payload.code ?? current?.code ?? "").trim(),
      type,
      scope,
      value: Number.isFinite(value) ? value : 0,
      start_date: normalizeDateString(payload.start_date ?? current?.start_date ?? ""),
      end_date: normalizeDateString(payload.end_date ?? current?.end_date ?? ""),
      min_purchase_cents: centsFromValue(payload.min_purchase_amount ?? current?.min_purchase_cents / 100 ?? 0),
      max_discount_cents: centsFromValue(payload.max_discount_amount ?? current?.max_discount_cents / 100 ?? 0),
      applicable_categories: listToValue(payload.applicable_categories ?? current?.applicable_categories ?? ""),
      applicable_product_ids: listToValue(payload.applicable_product_ids ?? current?.applicable_product_ids ?? ""),
      applicable_member_tiers: applicableMemberTiers,
      member_tier: applicableMemberTiers,
      promo_type: String(payload.promo_type ?? current?.promo_type ?? "").trim(),
      promo_buy_qty: toNonNegativeInt(payload.promo_buy_qty ?? current?.promo_buy_qty ?? 0, 0),
      promo_free_qty: toNonNegativeInt(payload.promo_free_qty ?? current?.promo_free_qty ?? 0, 0),
      promo_discount_percent: Number(payload.promo_discount_percent ?? current?.promo_discount_percent ?? 0),
      promo_spend_cents: centsFromValue(payload.promo_spend_amount ?? current?.promo_spend_cents / 100 ?? 0),
      promo_discount_cents: centsFromValue(payload.promo_discount_amount ?? current?.promo_discount_cents / 100 ?? 0),
      is_active: payload.is_active === false ? 0 : payload.is_active === true ? 1 : current ? current.is_active : 1,
      created_by: current?.created_by || String(actor || "admin"),
      created_at: current?.created_at || timestamp,
      updated_at: timestamp
    };
  }

  _isDiscountActiveNow(discount, value = new Date()) {
    if (!discount || !discount.is_active) {
      return false;
    }

    const today = dateKeyLocal(value);
    const start = normalizeDateString(discount.start_date);
    const end = normalizeDateString(discount.end_date);
    if (start && today < start) {
      return false;
    }
    if (end && today > end) {
      return false;
    }
    return true;
  }

  getDiscounts({ includeInactive = false, activeNow = false } = {}) {
    return this.discounts
      .filter((discount) => {
        if (activeNow) {
          return this._isDiscountActiveNow(discount);
        }
        return includeInactive || discount.is_active;
      })
      .slice()
      .sort((a, b) => Number(a.id) - Number(b.id))
      .map((discount) => ({ ...discount }));
  }

  getDiscountById(discountId) {
    const id = Number(discountId);
    const discount = this.discounts.find((entry) => Number(entry.id) === id);
    return discount ? { ...discount } : null;
  }

  addDiscount(payload, actor = "admin") {
    const discount = this._normalizeDiscountPayload(payload, null, actor);
    this.discounts.push(discount);
    this._logAudit({
      action: "discount_add",
      module: "Discount Management",
      changedBy: actor,
      description: `Created discount ${discount.name}`,
      details: { after: discount }
    });
    this._save();
    return { ...discount };
  }

  updateDiscount(discountId, payload, actor = "admin") {
    const id = Number(discountId);
    const idx = this.discounts.findIndex((discount) => Number(discount.id) === id);
    if (idx < 0) {
      const err = new Error("Discount not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const before = { ...this.discounts[idx] };
    const updated = this._normalizeDiscountPayload(payload, this.discounts[idx], actor);
    this.discounts[idx] = updated;
    this._logAudit({
      action: "discount_edit",
      module: "Discount Management",
      changedBy: actor,
      description: `Updated discount ${updated.name}`,
      details: { before, after: updated }
    });
    this._save();
    return { ...updated };
  }

  removeDiscount(discountId, actor = "admin") {
    const id = Number(discountId);
    const idx = this.discounts.findIndex((discount) => Number(discount.id) === id);
    if (idx < 0) {
      const err = new Error("Discount not found.");
      err.code = "NOT_FOUND";
      throw err;
    }
    const [removed] = this.discounts.splice(idx, 1);
    this._logAudit({
      action: "discount_delete",
      module: "Discount Management",
      changedBy: actor,
      description: `Deleted discount ${removed.name}`,
      details: { before: removed }
    });
    this._save();
    return { ...removed };
  }

  getDiscountUsage({ limit = 500 } = {}) {
    const max = Math.min(2000, Math.max(1, toNonNegativeInt(limit, 500) || 500));
    return this.discountUsage
      .slice()
      .sort((a, b) => Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((row) => ({ ...row }));
  }

  getPaymentByOrderId(orderId) {
    const rows = this.payments
      .filter((payment) => Number(payment.order_id) === Number(orderId))
      .slice()
      .sort((a, b) => Number(a.id) - Number(b.id));
    if (!rows.length) {
      return null;
    }
    const latest = rows[rows.length - 1];
    return {
      ...latest,
      amount_cents: rows.reduce((sum, payment) => sum + toNonNegativeInt(payment.amount_cents, 0), 0)
    };
  }

  getPaymentsByOrderId(orderId) {
    return this.payments
      .filter((payment) => Number(payment.order_id) === Number(orderId))
      .slice()
      .sort((a, b) => Number(a.id) - Number(b.id))
      .map((payment) => ({ ...payment }));
  }

  getTodaySummary(targetDate = "") {
    const dateKey = String(targetDate || "").trim() || dateKeyLocal(new Date());
    const orders = this.orders.filter((order) => dateKeyLocal(order.created_at) === dateKey && order.status === "paid");
    const orderIds = new Set(orders.map((order) => Number(order.id)));
    const totalCents = orders.reduce((sum, order) => sum + toNonNegativeInt(order.total_cents, 0), 0);
    const subtotalCents = orders.reduce((sum, order) => sum + toNonNegativeInt(order.subtotal_cents, 0), 0);
    const totalRoundingCents = orders.reduce((sum, order) => sum + toInt(order.rounding_cents, 0), 0);

    let cogsCents = 0;
    for (const item of this.orderItems) {
      if (!orderIds.has(Number(item.order_id))) {
        continue;
      }
      const quantity = toNonNegativeInt(item.quantity, 0);
      const unitCostCents = toNonNegativeInt(item.unit_cost_cents, 0);
      cogsCents += quantity * unitCostCents;
    }
    const profitCents = totalCents - cogsCents;

    const paymentMethodByOrderId = new Map();
    const payments = this.payments
      .slice()
      .sort((a, b) => Number(a.id) - Number(b.id));
    for (const payment of payments) {
      const orderId = Number(payment.order_id);
      if (!paymentMethodByOrderId.has(orderId)) {
        paymentMethodByOrderId.set(orderId, normalizePaymentMethod(payment.payment_method, "other"));
      }
    }

    const byMethod = {};
    for (const order of orders) {
      const key = paymentMethodByOrderId.get(Number(order.id)) || "other";
      byMethod[key] = (byMethod[key] || 0) + toNonNegativeInt(order.total_cents, 0);
    }

    const customerKeys = new Set(orders.map((order) => {
      const memberId = toNonNegativeInt(order.member_id, 0);
      const phone = String(order.phone || "").trim().toLowerCase();
      const name = String(order.customer_name || "").trim().toLowerCase();
      if (memberId > 0) return `member:${memberId}`;
      if (phone) return `phone:${phone}`;
      if (name) return `name:${name}`;
      return `walk-in:${order.id}`;
    }));

    return {
      date: dateKey,
      order_count: orders.length,
      customer_count: customerKeys.size,
      subtotal_cents: subtotalCents,
      total_cents: totalCents,
      cogs_cents: cogsCents,
      profit_cents: profitCents,
      rounding_cents: totalRoundingCents,
      payment_methods_cents: byMethod
    };
  }

  getSalesBreakdown(targetDate = "", limit = 20) {
    const dateKey = String(targetDate || "").trim() || dateKeyLocal(new Date());
    const numericLimit = Math.max(1, Math.min(200, toNonNegativeInt(limit, 20)));
    const paidOrders = this.orders.filter((order) => dateKeyLocal(order.created_at) === dateKey && order.status === "paid");
    const paidOrderIds = new Set(paidOrders.map((order) => Number(order.id)));
    const productById = new Map(this.products.map((product) => [Number(product.id), product]));

    const itemsMap = new Map();
    const categoriesMap = new Map();
    let totalUnits = 0;

    for (const row of this.orderItems) {
      const orderId = Number(row.order_id);
      if (!paidOrderIds.has(orderId)) {
        continue;
      }

      const quantity = toNonNegativeInt(row.quantity, 0);
      if (quantity <= 0 || toNonNegativeInt(row.product_id, 0) <= 0) {
        continue;
      }

      totalUnits += quantity;
      const itemName = String(row.product_name || "").trim() || "Unknown Item";
      const existingItem = itemsMap.get(itemName) || { item_name: itemName, quantity: 0 };
      existingItem.quantity += quantity;
      itemsMap.set(itemName, existingItem);

      const product = productById.get(Number(row.product_id));
      const categoryName = String(product?.category || "").trim() || "Uncategorized";
      const existingCategory = categoriesMap.get(categoryName) || { category: categoryName, quantity: 0 };
      existingCategory.quantity += quantity;
      categoriesMap.set(categoryName, existingCategory);
    }

    const byItem = Array.from(itemsMap.values())
      .sort((a, b) => b.quantity - a.quantity || a.item_name.localeCompare(b.item_name))
      .slice(0, numericLimit);
    const byCategory = Array.from(categoriesMap.values())
      .sort((a, b) => b.quantity - a.quantity || a.category.localeCompare(b.category))
      .slice(0, numericLimit);

    return {
      date: dateKey,
      order_count: paidOrders.length,
      total_units: totalUnits,
      distinct_items: itemsMap.size,
      distinct_categories: categoriesMap.size,
      by_item: byItem,
      by_category: byCategory
    };
  }

  getSoldItemsReport(filters = {}) {
    const range = reportDateRange({
      preset: filters.preset || "today",
      startDate: filters.start_date,
      endDate: filters.end_date
    });
    const normalizedSearch = String(filters.search || "").trim().toLowerCase();
    const normalizedCategory = String(filters.category || "").trim().toLowerCase();
    const normalizedBrand = String(filters.brand || "").trim().toLowerCase();
    const normalizedPaymentMethod = normalizePaymentMethod(filters.payment_method || "", "");
    const cashierFilter = String(filters.cashier || "").trim().toLowerCase();
    const reportStatuses = new Set(["paid", "unpaid", "partially_paid"]);
    const productById = new Map(this.products.map((product) => [Number(product.id), product]));
    const userById = new Map(this.users.map((user) => [Number(user.id), user]));
    const memberById = new Map(this.members.map((member) => [Number(member.id), member]));
    const orderItemsByOrderId = new Map();

    for (const item of this.orderItems) {
      const orderId = Number(item.order_id);
      if (!orderItemsByOrderId.has(orderId)) {
        orderItemsByOrderId.set(orderId, []);
      }
      orderItemsByOrderId.get(orderId).push(item);
    }

    const allOrdersInRange = this.orders
      .filter((order) => reportStatuses.has(normalizeOrderStatus(order.status, "")))
      .filter((order) => {
        const key = dateKeyLocal(order.created_at);
        return key >= range.start_date && key <= range.end_date;
      })
      .filter((order) => {
        if (!cashierFilter) {
          return true;
        }
        const user = userById.get(Number(order.user_id));
        return [
          order.user_id,
          user?.id,
          user?.username,
          user?.full_name,
          user?.role
        ].some((field) => String(field || "").toLowerCase() === cashierFilter);
      })
      .filter((order) => {
        if (!normalizedPaymentMethod) {
          return true;
        }
        const methods = this.getPaymentsByOrderId(order.id).map((payment) => normalizePaymentMethod(payment.payment_method, ""));
        if (normalizedPaymentMethod === "havent_paid") {
          return normalizeOrderStatus(order.status, "") === "unpaid" || methods.includes("havent_paid");
        }
        return methods.includes(normalizedPaymentMethod);
      });

    const itemRows = [];

    for (const order of allOrdersInRange) {
      const orderItems = (orderItemsByOrderId.get(Number(order.id)) || [])
        .filter((item) => toNonNegativeInt(item.product_id, 0) > 0)
        .slice()
        .sort((a, b) => Number(a.id) - Number(b.id));
      const subtotalCents = Math.max(
        1,
        orderItems.reduce((sum, item) => sum + toNonNegativeInt(item.line_total_cents, 0), 0)
      );
      let remainingDiscountCents = toNonNegativeInt(order.discount_cents, 0);
      const user = userById.get(Number(order.user_id));
      const member = memberById.get(Number(order.member_id));

      for (const [index, item] of orderItems.entries()) {
        const product = productById.get(Number(item.product_id));
        const quantity = toNonNegativeInt(item.quantity, 0);
        if (quantity <= 0) {
          continue;
        }

        const lineTotalCents = toNonNegativeInt(item.line_total_cents, 0);
        const discountCents =
          index === orderItems.length - 1
            ? remainingDiscountCents
            : Math.min(remainingDiscountCents, Math.round(toNonNegativeInt(order.discount_cents, 0) * lineTotalCents / subtotalCents));
        remainingDiscountCents = Math.max(0, remainingDiscountCents - discountCents);
        const category = String(product?.category || "Uncategorized").trim() || "Uncategorized";
        const brand = String(product?.brand || "").trim();
        const barcode = String(product?.barcode || "").trim();
        const sku = String(product?.sku || "").trim();
        const productName = String(item.product_name || product?.name || "Unknown Item").trim();
        const receiptNumber = String(order.order_number || "");
        const customer = member?.name || order.customer_name || "";

        const row = {
          id: `${order.id}-${item.id}`,
          order_id: order.id,
          order_item_id: item.id,
          created_at: order.created_at,
          date: dateKeyLocal(order.created_at),
          receipt_number: receiptNumber,
          invoice_number: String(order.invoice_number || ""),
          product_id: toNonNegativeInt(item.product_id, 0),
          product_name: productName,
          barcode,
          sku,
          category,
          brand,
          quantity,
          unit_price_cents: toNonNegativeInt(item.unit_price_cents, 0),
          line_total_cents: lineTotalCents,
          discount_cents: discountCents,
          net_total_cents: Math.max(0, lineTotalCents - discountCents),
          unit_cost_cents: toNonNegativeInt(item.unit_cost_cents, toNonNegativeInt(product?.cost_cents, 0)),
          remaining_stock: toInt(product?.quantity_on_hand, 0),
          reorder_level: toNonNegativeInt(product?.reorder_level, 0),
          low_stock: Boolean(product && toInt(product.quantity_on_hand, 0) <= toNonNegativeInt(product.reorder_level, 0)),
          cashier_id: toNonNegativeInt(order.user_id, 0),
          cashier: user ? user.full_name || user.username : "",
          customer,
          payment_status: normalizeOrderStatus(order.status, "paid")
        };

        if (normalizedCategory && category.toLowerCase() !== normalizedCategory) {
          continue;
        }
        if (normalizedBrand && brand.toLowerCase() !== normalizedBrand) {
          continue;
        }
        if (normalizedSearch) {
          const matches = [productName, barcode, sku, receiptNumber, row.invoice_number, customer].some((field) =>
            String(field || "").toLowerCase().includes(normalizedSearch)
          );
          if (!matches) {
            continue;
          }
        }

        itemRows.push(row);
      }
    }

    const includedOrderIds = new Set(itemRows.map((row) => Number(row.order_id)));
    const includedOrders = allOrdersInRange.filter((order) => includedOrderIds.has(Number(order.id)));
    const categories = new Map();
    const products = new Map();
    const staff = new Map();

    for (const row of itemRows) {
      const categoryKey = row.category || "Uncategorized";
      if (!categories.has(categoryKey)) {
        categories.set(categoryKey, {
          category: categoryKey,
          total_quantity: 0,
          revenue_cents: 0,
          discount_cents: 0,
          products: new Map()
        });
      }
      const category = categories.get(categoryKey);
      category.total_quantity += row.quantity;
      category.revenue_cents += row.net_total_cents;
      category.discount_cents += row.discount_cents;
      const categoryProduct = category.products.get(row.product_id) || {
        product_id: row.product_id,
        product_name: row.product_name,
        quantity: 0,
        revenue_cents: 0,
        discount_cents: 0
      };
      categoryProduct.quantity += row.quantity;
      categoryProduct.revenue_cents += row.net_total_cents;
      categoryProduct.discount_cents += row.discount_cents;
      category.products.set(row.product_id, categoryProduct);

      const productKey = row.product_id || row.product_name;
      const product = products.get(productKey) || {
        product_id: row.product_id,
        product_name: row.product_name,
        category: row.category,
        brand: row.brand,
        barcode: row.barcode,
        sku: row.sku,
        quantity: 0,
        revenue_cents: 0,
        discount_cents: 0,
        cost_cents: 0,
        remaining_stock: row.remaining_stock,
        reorder_level: row.reorder_level,
        low_stock: row.low_stock
      };
      product.quantity += row.quantity;
      product.revenue_cents += row.net_total_cents;
      product.discount_cents += row.discount_cents;
      product.cost_cents += row.unit_cost_cents * row.quantity;
      product.remaining_stock = row.remaining_stock;
      product.reorder_level = row.reorder_level;
      product.low_stock = row.low_stock;
      products.set(productKey, product);

      const staffKey = row.cashier_id || row.cashier || "unknown";
      const staffRow = staff.get(staffKey) || {
        cashier_id: row.cashier_id,
        cashier: row.cashier || "Unknown",
        order_ids: new Set(),
        items_sold: 0,
        revenue_cents: 0,
        discount_cents: 0
      };
      staffRow.order_ids.add(Number(row.order_id));
      staffRow.items_sold += row.quantity;
      staffRow.revenue_cents += row.net_total_cents;
      staffRow.discount_cents += row.discount_cents;
      staff.set(staffKey, staffRow);
    }

    const paymentSummary = {
      cash_cents: 0,
      qr_cents: 0,
      card_cents: 0,
      bank_transfer_cents: 0,
      e_wallet_cents: 0,
      unpaid_cents: 0,
      outstanding_cents: 0
    };

    for (const order of includedOrders) {
      const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
      const paidCents = toNonNegativeInt(this.getPaymentByOrderId(order.id)?.amount_cents, 0);
      const outstandingCents = Math.max(0, payableCents - paidCents);
      if (normalizeOrderStatus(order.status, "") === "unpaid") {
        paymentSummary.unpaid_cents += outstandingCents;
      }
      paymentSummary.outstanding_cents += outstandingCents;

      let remainingPaymentCents = payableCents;
      for (const payment of this.getPaymentsByOrderId(order.id)) {
        const amountCents = Math.min(toNonNegativeInt(payment.amount_cents, 0), remainingPaymentCents);
        if (amountCents <= 0) {
          continue;
        }
        remainingPaymentCents = Math.max(0, remainingPaymentCents - amountCents);
        const method = normalizePaymentMethod(payment.payment_method, "");
        if (method === "cash") {
          paymentSummary.cash_cents += amountCents;
        } else if (method === "duitnow_qr") {
          paymentSummary.qr_cents += amountCents;
        } else if (method === "credit_card" || method === "debit_card") {
          paymentSummary.card_cents += amountCents;
        } else if (method === "bank_transfer") {
          paymentSummary.bank_transfer_cents += amountCents;
        } else if (method === "e_wallet") {
          paymentSummary.e_wallet_cents += amountCents;
        }
      }
    }

    const productSummary = Array.from(products.values())
      .map((product) => ({
        ...product,
        average_selling_price_cents: product.quantity > 0 ? Math.round(product.revenue_cents / product.quantity) : 0,
        profit_cents: product.revenue_cents - product.cost_cents,
        margin_percent:
          product.revenue_cents > 0
            ? Number((((product.revenue_cents - product.cost_cents) / product.revenue_cents) * 100).toFixed(2))
            : 0
      }))
      .sort((a, b) => b.quantity - a.quantity || b.revenue_cents - a.revenue_cents || a.product_name.localeCompare(b.product_name));

    const categorySummary = Array.from(categories.values())
      .map((category) => ({
        category: category.category,
        total_quantity: category.total_quantity,
        revenue_cents: category.revenue_cents,
        discount_cents: category.discount_cents,
        products: Array.from(category.products.values())
          .sort((a, b) => b.quantity - a.quantity || a.product_name.localeCompare(b.product_name))
      }))
      .sort((a, b) => b.revenue_cents - a.revenue_cents || a.category.localeCompare(b.category));

    const staffSummary = Array.from(staff.values())
      .map((row) => ({
        cashier_id: row.cashier_id,
        cashier: row.cashier,
        orders_processed: row.order_ids.size,
        items_sold: row.items_sold,
        revenue_cents: row.revenue_cents,
        discount_cents: row.discount_cents
      }))
      .sort((a, b) => b.revenue_cents - a.revenue_cents || a.cashier.localeCompare(b.cashier));

    return {
      range,
      generated_at: nowISO(),
      filters: {
        preset: range.preset,
        start_date: range.start_date,
        end_date: range.end_date,
        search: filters.search || "",
        category: filters.category || "",
        brand: filters.brand || "",
        cashier: filters.cashier || "",
        payment_method: filters.payment_method || "",
        categories: Array.from(new Set(this.products.map((product) => String(product.category || "").trim()).filter(Boolean))).sort(),
        brands: Array.from(new Set(this.products.map((product) => String(product.brand || "").trim()).filter(Boolean))).sort(),
        cashiers: this.users
          .filter((user) => user.is_active)
          .map((user) => ({
            id: user.id,
            username: user.username,
            full_name: user.full_name,
            role: user.role
          }))
          .sort((a, b) => String(a.full_name || a.username).localeCompare(String(b.full_name || b.username)))
      },
      totals: {
        order_count: includedOrders.length,
        item_count: itemRows.length,
        quantity_sold: itemRows.reduce((sum, row) => sum + row.quantity, 0),
        revenue_cents: itemRows.reduce((sum, row) => sum + row.net_total_cents, 0),
        discount_cents: itemRows.reduce((sum, row) => sum + row.discount_cents, 0)
      },
      item_rows: itemRows.sort((a, b) => String(b.created_at).localeCompare(String(a.created_at))),
      category_summary: categorySummary,
      product_summary: productSummary,
      top_products: productSummary.slice(0, 10).map((product, index) => ({
        rank: index + 1,
        product_id: product.product_id,
        product_name: product.product_name,
        quantity: product.quantity,
        revenue_cents: product.revenue_cents
      })),
      low_stock: productSummary.filter((product) => product.low_stock),
      payment_summary: paymentSummary,
      staff_summary: staffSummary
    };
  }

  addProduct(payload, actor = "admin") {
    const sku = String(payload.sku || "").trim();
    const name = String(payload.name || "").trim();
    const category = String(payload.category || "").trim() || "Uncategorized";
    const subcategory = String(payload.subcategory || "").trim() || "General";
    const barcode = String(payload.barcode || "").trim();
    const brand = String(payload.brand || "").trim();
    const supplier = String(payload.supplier || "").trim();
    const description = String(payload.description || "").trim();
    const templateType = String(payload.template_type || "custom").trim().toLowerCase() || "custom";
    const attributes = payload.attributes && typeof payload.attributes === "object" ? payload.attributes : {};
    const unitPriceCents = centsFromValue(payload.unit_price);
    const costCents = centsFromValue(payload.cost);
    const quantity = toNonNegativeInt(payload.quantity_on_hand, 0);
    const reorder = toNonNegativeInt(payload.reorder_level, 0);
    const isActive = payload.is_active === false ? 0 : 1;

    if (!sku || !name) {
      const err = new Error("SKU and name are required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const duplicate = this.products.find((product) => product.sku.toLowerCase() === sku.toLowerCase());
    if (duplicate) {
      const err = new Error("SKU already exists.");
      err.code = "CONFLICT";
      throw err;
    }

    const timestamp = nowISO();
    const product = {
      id: this._nextId("product_id"),
      sku,
      name,
      category,
      subcategory,
      barcode,
      brand,
      supplier,
      description,
      template_type: templateType,
      attributes_json: JSON.stringify(attributes),
      unit_price_cents: unitPriceCents,
      cost_cents: costCents,
      quantity_on_hand: quantity,
      reorder_level: reorder,
      is_active: isActive,
      created_at: timestamp,
      updated_at: timestamp
    };

    this.products.push(product);
    this._logEdit({
      productId: product.id,
      action: "add",
      changedBy: actor,
      details: JSON.stringify({ after: product })
    });
    if (quantity > 0) {
      this._logInventoryRecord({
        productId: product.id,
        action: "item_create",
        quantityDelta: quantity,
        beforeQty: 0,
        afterQty: quantity,
        reason: "Initial stock",
        changedBy: actor
      });
    }

    this._save();
    return { ...product };
  }

  updateProduct(productId, payload, actor = "admin") {
    const id = Number(productId);
    const idx = this.products.findIndex((product) => product.id === id);
    if (idx < 0) {
      const err = new Error("Product not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const current = this.products[idx];
    const sku = String(payload.sku || "").trim();
    const name = String(payload.name || "").trim();

    if (!sku || !name) {
      const err = new Error("SKU and name are required.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const duplicate = this.products.find(
      (product) => product.id !== id && product.sku.toLowerCase() === sku.toLowerCase()
    );
    if (duplicate) {
      const err = new Error("SKU already exists.");
      err.code = "CONFLICT";
      throw err;
    }

    const before = { ...current };
    const updated = {
      ...current,
      sku,
      name,
      category: String(payload.category || "").trim() || "Uncategorized",
      subcategory: String(payload.subcategory || "").trim() || "General",
      barcode: String(payload.barcode || "").trim(),
      brand: String(payload.brand || "").trim(),
      supplier: String(payload.supplier || "").trim(),
      description: String(payload.description || "").trim(),
      template_type: String(payload.template_type || current.template_type || "custom").trim().toLowerCase() || "custom",
      attributes_json:
        payload.attributes && typeof payload.attributes === "object"
          ? JSON.stringify(payload.attributes)
          : String(current.attributes_json || "{}"),
      unit_price_cents: centsFromValue(payload.unit_price),
      cost_cents: Object.hasOwn(payload, "cost") ? centsFromValue(payload.cost) : toNonNegativeInt(current.cost_cents, 0),
      quantity_on_hand: toNonNegativeInt(payload.quantity_on_hand, 0),
      reorder_level: toNonNegativeInt(payload.reorder_level, 0),
      is_active: payload.is_active === false ? 0 : 1,
      updated_at: nowISO()
    };

    this.products[idx] = updated;
    this._logEdit({
      productId: updated.id,
      action: "edit",
      changedBy: actor,
      details: JSON.stringify({ before, after: updated })
    });
    if (toNonNegativeInt(before.quantity_on_hand, 0) !== toNonNegativeInt(updated.quantity_on_hand, 0)) {
      this._logInventoryRecord({
        productId: updated.id,
        action: "stock_edit",
        quantityDelta: toNonNegativeInt(updated.quantity_on_hand, 0) - toNonNegativeInt(before.quantity_on_hand, 0),
        beforeQty: toNonNegativeInt(before.quantity_on_hand, 0),
        afterQty: toNonNegativeInt(updated.quantity_on_hand, 0),
        reason: String(payload.stock_reason || "Item edit"),
        changedBy: actor
      });
    }

    this._save();
    return { ...updated };
  }

  removeProduct(productId, actor = "admin") {
    const id = Number(productId);
    const idx = this.products.findIndex((product) => product.id === id);
    if (idx < 0) {
      const err = new Error("Product not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    if (!this.products[idx].is_active) {
      const err = new Error("Product is already removed from sale.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const before = { ...this.products[idx] };
    this.products[idx].is_active = 0;
    this.products[idx].updated_at = nowISO();

    this._logEdit({
      productId: this.products[idx].id,
      action: "remove",
      changedBy: actor,
      details: JSON.stringify({ before, after: this.products[idx] })
    });

    this._save();
    return { ...this.products[idx] };
  }

  _nextOrderSequenceForDate(value = new Date()) {
    const date = value instanceof Date ? value : new Date(value);
    const dateKey = dateKeyLocal(date);
    const datePart = dateKey.replace(/-/g, "");
    let maxSequence = 0;
    let sameDateCount = 0;

    for (const order of this.orders) {
      if (dateKeyLocal(order.created_at) !== dateKey) {
        continue;
      }

      sameDateCount += 1;
      const orderNo = String(order.order_number || "").trim();
      const match = orderNo.match(/^ORD-(\d{8})-(\d+)$/i);
      if (!match) {
        continue;
      }

      if (match[1] !== datePart) {
        continue;
      }

      const seq = toNonNegativeInt(match[2], 0);
      if (seq > maxSequence) {
        maxSequence = seq;
      }
    }

    if (maxSequence > 0) {
      return maxSequence + 1;
    }

    // Backward compatibility for older order number formats on same date.
    return sameDateCount + 1;
  }

  _invoiceSequenceState() {
    const parsed = safeJsonParse(this.getSetting("invoice_sequence_json", "{}"), {});
    return parsed && typeof parsed === "object" && !Array.isArray(parsed) ? parsed : {};
  }

  _invoiceSequenceKey(settings, value = new Date()) {
    const normalized = normalizeInvoiceSettings(settings);
    const token = invoiceDateToken(value, normalized.date_format);
    const period = invoicePeriodKey(value, normalized.reset);
    return [
      normalized.prefix.toLowerCase(),
      normalized.date_format.toLowerCase(),
      normalized.reset,
      period,
      token
    ].join(":");
  }

  _invoiceNumberExists(invoiceNumber, excludeOrderId = 0) {
    const normalized = String(invoiceNumber || "").trim().toLowerCase();
    if (!normalized) {
      return false;
    }
    return this.orders.some((order) =>
      Number(order.id) !== Number(excludeOrderId) &&
      String(order.invoice_number || "").trim().toLowerCase() === normalized
    );
  }

  previewNextInvoiceNumber(settings = {}, value = new Date(), excludeOrderId = 0) {
    const normalized = normalizeInvoiceSettings(settings);
    const at = value instanceof Date ? value : new Date(value);
    const token = invoiceDateToken(at, normalized.date_format);
    const sequenceKey = this._invoiceSequenceKey(normalized, at);
    const sequenceState = this._invoiceSequenceState();
    let sequence = toNonNegativeInt(sequenceState[sequenceKey]?.last_sequence, 0) + 1;
    let invoiceNumber = invoiceNumberFromIdentity(normalized, token, sequence);

    while (this._invoiceNumberExists(invoiceNumber, excludeOrderId)) {
      sequence += 1;
      invoiceNumber = invoiceNumberFromIdentity(normalized, token, sequence);
    }

    return {
      invoice_number: invoiceNumber,
      sequence,
      sequence_key: sequenceKey,
      date_token: token,
      settings: normalized
    };
  }

  ensureOrderInvoiceNumber(orderId, settings = {}, value = new Date(), actor = "system") {
    const id = Number(orderId);
    const orderIdx = this.orders.findIndex((entry) => Number(entry.id) === id);
    if (orderIdx < 0) {
      const err = new Error("Order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const order = this.orders[orderIdx];
    if (String(order.invoice_number || "").trim()) {
      return this.getOrderById(id);
    }

    const generatedAt = value instanceof Date ? value.toISOString() : String(value || nowISO());
    const invoiceSettings = normalizeInvoiceSettings(settings);
    const next = this.previewNextInvoiceNumber(invoiceSettings, generatedAt, id);
    order.invoice_number = next.invoice_number;
    order.invoice_generated_at = generatedAt;

    const sequenceState = this._invoiceSequenceState();
    const currentSequence = toNonNegativeInt(sequenceState[next.sequence_key]?.last_sequence, 0);
    sequenceState[next.sequence_key] = {
      last_sequence: Math.max(currentSequence, next.sequence),
      last_invoice_number: next.invoice_number,
      updated_at: generatedAt
    };
    this._setSettingValue("invoice_sequence_json", JSON.stringify(sequenceState));
    this._setSettingValue("last_invoice_number", next.invoice_number);

    this.orders[orderIdx] = order;
    this._logAudit({
      action: "invoice_number_assigned",
      module: "Receipt Management",
      changedBy: actor,
      description: `Assigned invoice number ${next.invoice_number}`,
      details: {
        order_id: id,
        invoice_number: next.invoice_number,
        sequence_key: next.sequence_key,
        sequence: next.sequence
      }
    });
    this._save();
    return this.getOrderById(id);
  }

  recordInvoiceExport(orderId, payload = {}, actor = "system") {
    const id = Number(orderId);
    const orderIdx = this.orders.findIndex((entry) => Number(entry.id) === id);
    if (orderIdx < 0) {
      const err = new Error("Order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const generatedAt = String(payload.generated_at || nowISO());
    const order = this.orders[orderIdx];
    let invoiceNumber = String(order.invoice_number || "").trim();
    let sequence = toNonNegativeInt(payload.sequence, 0);
    let sequenceKey = String(payload.sequence_key || "").trim();
    const invoiceSettings = normalizeInvoiceSettings(payload.invoice_settings || {});

    if (!invoiceNumber) {
      const requestedNumber = String(payload.invoice_number || "").trim();
      if (requestedNumber) {
        if (this._invoiceNumberExists(requestedNumber, id)) {
          const err = new Error(`Invoice number already exists: ${requestedNumber}`);
          err.code = "CONFLICT";
          throw err;
        }
        invoiceNumber = requestedNumber;
        sequenceKey = sequenceKey || this._invoiceSequenceKey(invoiceSettings, generatedAt);
      } else {
        const next = this.previewNextInvoiceNumber(invoiceSettings, generatedAt, id);
        invoiceNumber = next.invoice_number;
        sequence = next.sequence;
        sequenceKey = next.sequence_key;
      }

      order.invoice_number = invoiceNumber;
      order.invoice_generated_at = generatedAt;

      if (sequenceKey) {
        const sequenceState = this._invoiceSequenceState();
        const currentSequence = toNonNegativeInt(sequenceState[sequenceKey]?.last_sequence, 0);
        sequenceState[sequenceKey] = {
          last_sequence: Math.max(currentSequence, sequence),
          last_invoice_number: invoiceNumber,
          updated_at: generatedAt
        };
        this._setSettingValue("invoice_sequence_json", JSON.stringify(sequenceState));
        this._setSettingValue("last_invoice_number", invoiceNumber);
      }
    }

    order.invoice_pdf_file_name = String(payload.pdf_file_name || order.invoice_pdf_file_name || "").trim();
    order.invoice_pdf_file_path = String(payload.pdf_file_path || order.invoice_pdf_file_path || "").trim();
    order.invoice_pdf_generated_at = generatedAt;
    order.invoice_export_count = toNonNegativeInt(order.invoice_export_count, 0) + 1;
    order.invoice_last_exported_at = generatedAt;
    if (payload.print) {
      order.invoice_print_count = toNonNegativeInt(order.invoice_print_count, 0) + 1;
      order.invoice_last_printed_at = generatedAt;
    } else {
      order.invoice_print_count = toNonNegativeInt(order.invoice_print_count, 0);
      order.invoice_last_printed_at = String(order.invoice_last_printed_at || "");
    }

    this.orders[orderIdx] = order;
    this._logAudit({
      action: payload.print ? "invoice_printed" : "invoice_exported",
      module: "Receipt Management",
      changedBy: actor,
      description: `${payload.print ? "Printed" : "Exported"} invoice ${invoiceNumber}`,
      details: {
        order_id: id,
        invoice_number: invoiceNumber,
        pdf_file_name: order.invoice_pdf_file_name,
        pdf_file_path: order.invoice_pdf_file_path
      }
    });
    this._save();
    return this.getOrderById(id);
  }

  _discountEligibleItems(discount, normalizedItems) {
    const productIds = new Set(idListFromValue(discount.applicable_product_ids));
    const categories = new Set(listFromValue(discount.applicable_categories).map((item) => item.toLowerCase()));
    const type = String(discount.type || "").toLowerCase();

    return normalizedItems.filter((item) => {
      const product = this.products.find((entry) => Number(entry.id) === Number(item.product_id));
      if (!product) {
        return false;
      }

      if (productIds.size > 0 && !productIds.has(Number(item.product_id))) {
        return false;
      }

      if (categories.size > 0 && !categories.has(String(product.category || "").trim().toLowerCase())) {
        return false;
      }

      if (type === "product" && productIds.size === 0) {
        return false;
      }

      if (type === "category" && categories.size === 0) {
        return false;
      }

      return true;
    });
  }

  _calculateDiscountRule(discount, normalizedItems, subtotalCents, member = null) {
    if (!discount || !this._isDiscountActiveNow(discount)) {
      return { discount_cents: 0, discount_label: "", discount: null };
    }

    const memberTiers = new Set(listFromValue(discount.applicable_member_tiers || discount.member_tier).map((item) => item.toLowerCase()));
    if (memberTiers.size > 0) {
      const memberTier = String(member?.tier || "").trim().toLowerCase();
      if (!memberTier || !memberTiers.has(memberTier)) {
        return { discount_cents: 0, discount_label: discount.name, discount };
      }
    }

    if (toNonNegativeInt(discount.min_purchase_cents, 0) > subtotalCents) {
      return { discount_cents: 0, discount_label: discount.name, discount };
    }

    const type = String(discount.type || "").toLowerCase();
    const value = Number(discount.value || 0);
    const eligibleItems = this._discountEligibleItems(discount, normalizedItems);
    const eligibleSubtotalCents = eligibleItems.reduce((sum, item) => sum + toNonNegativeInt(item.line_total_cents, 0), 0);
    let discountCents = 0;

    if (type === "percentage") {
      discountCents = Math.round(subtotalCents * Math.max(0, value) / 100);
    } else if (type === "fixed") {
      discountCents = centsFromValue(value);
    } else if (type === "fixed_per_unit") {
      for (const item of eligibleItems.length ? eligibleItems : normalizedItems) {
        discountCents += toNonNegativeInt(item.quantity, 0) * centsFromValue(value);
      }
    } else if (type === "member" && member) {
      discountCents = Math.round(subtotalCents * Math.max(0, value) / 100);
    } else if (type === "product") {
      discountCents = Math.round(eligibleSubtotalCents * Math.max(0, value) / 100);
    } else if (type === "category") {
      discountCents = Math.round(eligibleSubtotalCents * Math.max(0, value) / 100);
    } else if (type === "promotion") {
      const promoType = String(discount.promo_type || "").toLowerCase();
      if (promoType === "buy_x_free_y") {
        const buyQty = Math.max(1, toNonNegativeInt(discount.promo_buy_qty, 0));
        const freeQty = Math.max(1, toNonNegativeInt(discount.promo_free_qty, 0));
        for (const item of eligibleItems.length ? eligibleItems : normalizedItems) {
          const groupSize = buyQty + freeQty;
          const freeUnits = Math.floor(toNonNegativeInt(item.quantity, 0) / groupSize) * freeQty;
          discountCents += freeUnits * toNonNegativeInt(item.unit_price_cents, 0);
        }
      } else if (promoType === "buy_x_get_percent") {
        const buyQty = Math.max(1, toNonNegativeInt(discount.promo_buy_qty, 0));
        const percent = Number(discount.promo_discount_percent || value || 0);
        const targetItems = eligibleItems.length ? eligibleItems : normalizedItems;
        const targetQty = targetItems.reduce((sum, item) => sum + toNonNegativeInt(item.quantity, 0), 0);
        const targetSubtotal = targetItems.reduce((sum, item) => sum + toNonNegativeInt(item.line_total_cents, 0), 0);
        if (targetQty >= buyQty) {
          discountCents = Math.round(targetSubtotal * Math.max(0, percent) / 100);
        }
      } else if (promoType === "spend_get_fixed") {
        const spendCents = toNonNegativeInt(discount.promo_spend_cents, 0) || toNonNegativeInt(discount.min_purchase_cents, 0);
        if (subtotalCents >= spendCents) {
          discountCents = toNonNegativeInt(discount.promo_discount_cents, 0) || centsFromValue(value);
        }
      }
    }

    const capCents = toNonNegativeInt(discount.max_discount_cents, 0);
    if (capCents > 0) {
      discountCents = Math.min(discountCents, capCents);
    }

    discountCents = Math.min(subtotalCents, Math.max(0, toNonNegativeInt(discountCents, 0)));
    const label = discount.code ? `${discount.name} (${discount.code})` : discount.name;
    return { discount_cents: discountCents, discount_label: label, discount };
  }

  _customerKeyPart(value) {
    return String(value || "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  _customerIdentityFromPayload(payload = {}, fallbackOrderId = 0) {
    const memberId = toNonNegativeInt(payload.member_id, 0);
    const member = memberId ? this.getMemberById(memberId) : null;
    const customerName = String(payload.customer_name || payload.name || member?.name || "").trim() || "Walk-in";
    const phone = String(payload.phone || member?.phone || "").trim();
    let customerKey = String(payload.customer_key || "").trim();

    if (!customerKey) {
      if (member?.id || memberId) {
        customerKey = `member:${member?.id || memberId}`;
      } else if (phone) {
        customerKey = `phone:${this._customerKeyPart(phone)}`;
      } else if (customerName && customerName.toLowerCase() !== "walk-in") {
        customerKey = `name:${this._customerKeyPart(customerName)}`;
      } else {
        customerKey = `walk-in:${fallbackOrderId || "unknown"}`;
      }
    }

    return {
      customer_key: customerKey,
      customer_name: customerName,
      phone,
      member_id: member?.id || memberId
    };
  }

  _customerIdentityFromOrder(order = {}) {
    const member = this.getMemberById(order.member_id);
    return this._customerIdentityFromPayload(
      {
        member_id: order.member_id,
        customer_name: member?.name || order.customer_name,
        phone: member?.phone || order.phone
      },
      order.id
    );
  }

  _hasCustomerIdentity(payload = {}) {
    const identity = this._customerIdentityFromPayload(payload);
    return identity.customer_key && identity.customer_key !== "walk-in:unknown";
  }

  _recordLedgerEntry({ identity, orderId = 0, receiptNumber = "", description = "", debitCents = 0, creditCents = 0, actor = "system", createdAt = nowISO() }) {
    const debit = toNonNegativeInt(debitCents, 0);
    const credit = toNonNegativeInt(creditCents, 0);
    if (debit <= 0 && credit <= 0) {
      return null;
    }

    const row = {
      id: this._nextId("customer_ledger_id"),
      customer_key: identity.customer_key,
      customer_name: identity.customer_name || "Walk-in",
      phone: identity.phone || "",
      member_id: toNonNegativeInt(identity.member_id, 0),
      order_id: toNonNegativeInt(orderId, 0),
      receipt_number: String(receiptNumber || "").trim(),
      description: String(description || "").trim() || (debit > 0 ? "Customer Payment" : "Outstanding Created"),
      debit_cents: debit,
      credit_cents: credit,
      created_at: createdAt,
      created_by: String(actor || "system")
    };
    this.customerLedger.push(row);
    return row;
  }

  _orderOutstandingInvoice(orderId) {
    return this.outstandingInvoices.find((invoice) => Number(invoice.order_id) === Number(orderId)) || null;
  }

  _invoiceDueDate(invoice) {
    const createdAt = String(invoice?.created_at || nowISO());
    return dateKeyLocal(new Date(new Date(createdAt).getTime() + 7 * 24 * 60 * 60 * 1000));
  }

  _syncOrderStatusFromInvoice(invoice) {
    if (!invoice || !invoice.order_id) {
      return;
    }

    const order = this.orders.find((entry) => Number(entry.id) === Number(invoice.order_id));
    if (!order) {
      return;
    }

    if (toNonNegativeInt(invoice.balance_cents, 0) <= 0) {
      order.status = "paid";
    } else if (toNonNegativeInt(invoice.paid_cents, 0) > 0) {
      order.status = "partially_paid";
    } else {
      order.status = "unpaid";
    }
  }

  _upsertOutstandingInvoiceForOrder(order, balanceCents, actor = "system") {
    const balance = toNonNegativeInt(balanceCents, 0);
    if (!order || !order.id || balance <= 0) {
      return null;
    }

    const identity = this._customerIdentityFromOrder(order);
    const timestamp = order.created_at || nowISO();
    let invoice = this._orderOutstandingInvoice(order.id);
    if (!invoice) {
      invoice = {
        id: this._nextId("outstanding_invoice_id"),
        order_id: order.id,
        receipt_number: order.order_number || "",
        customer_key: identity.customer_key,
        customer_name: identity.customer_name,
        phone: identity.phone,
        member_id: identity.member_id,
        description: String(order.notes || "Retail Sale").trim() || "Retail Sale",
        original_cents: balance,
        paid_cents: 0,
        balance_cents: balance,
        status: normalizeOrderStatus(order.status || "unpaid", "unpaid"),
        created_at: timestamp,
        updated_at: nowISO()
      };
      this.outstandingInvoices.push(invoice);
    } else {
      invoice.customer_key = identity.customer_key;
      invoice.customer_name = identity.customer_name;
      invoice.phone = identity.phone;
      invoice.member_id = identity.member_id;
      invoice.description = String(order.notes || invoice.description || "Retail Sale").trim() || "Retail Sale";
      invoice.original_cents = Math.max(toNonNegativeInt(invoice.original_cents, 0), balance + toNonNegativeInt(invoice.paid_cents, 0));
      invoice.balance_cents = balance;
      invoice.status = toNonNegativeInt(invoice.paid_cents, 0) > 0 ? "partially_paid" : "unpaid";
      invoice.updated_at = nowISO();
    }

    const hasCreditEntry = this.customerLedger.some(
      (entry) => Number(entry.order_id) === Number(order.id) && toNonNegativeInt(entry.credit_cents, 0) > 0
    );
    if (!hasCreditEntry) {
      this._recordLedgerEntry({
        identity,
        orderId: order.id,
        receiptNumber: order.order_number,
        description: `Invoice ${order.order_number} Outstanding`,
        creditCents: balance,
        actor,
        createdAt: timestamp
      });
    }

    return invoice;
  }

  _applyCustomerPayment({ identity, amountCents, orderId = 0, targetOrderId = 0, receiptNumber = "", description = "Outstanding Payment", actor = "system", createdAt = nowISO() }) {
    const amount = toNonNegativeInt(amountCents, 0);
    if (!identity?.customer_key || amount <= 0) {
      return { applied_cents: 0, remaining_cents: amount };
    }

    this._recordLedgerEntry({
      identity,
      orderId,
      receiptNumber,
      description,
      debitCents: amount,
      actor,
      createdAt
    });

    let remaining = amount;
    const invoices = this.outstandingInvoices
      .filter((invoice) => invoice.customer_key === identity.customer_key && toNonNegativeInt(invoice.balance_cents, 0) > 0)
      .sort((a, b) => {
        const aIsTarget = Number(targetOrderId || 0) > 0 && Number(a.order_id) === Number(targetOrderId);
        const bIsTarget = Number(targetOrderId || 0) > 0 && Number(b.order_id) === Number(targetOrderId);
        if (aIsTarget !== bIsTarget) return aIsTarget ? -1 : 1;
        return String(a.created_at || "").localeCompare(String(b.created_at || "")) || Number(a.id) - Number(b.id);
      });

    for (const invoice of invoices) {
      if (remaining <= 0) {
        break;
      }

      const balance = toNonNegativeInt(invoice.balance_cents, 0);
      const applied = Math.min(balance, remaining);
      invoice.paid_cents = toNonNegativeInt(invoice.paid_cents, 0) + applied;
      invoice.balance_cents = balance - applied;
      invoice.status = invoice.balance_cents <= 0 ? "paid" : "partially_paid";
      invoice.updated_at = nowISO();
      this._syncOrderStatusFromInvoice(invoice);
      remaining -= applied;
    }

    return {
      applied_cents: amount - remaining,
      remaining_cents: remaining
    };
  }

  _ledgerRowsForCustomer(customerKey = "") {
    const normalizedKey = String(customerKey || "").trim();
    return this.customerLedger
      .filter((entry) => !normalizedKey || entry.customer_key === normalizedKey)
      .slice()
      .sort((a, b) => String(a.created_at || "").localeCompare(String(b.created_at || "")) || Number(a.id) - Number(b.id));
  }

  getCustomerOutstanding(payload = {}) {
    if (!this._hasCustomerIdentity(payload)) {
      const err = new Error("Select a customer or member before checking outstanding balance.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const identity = this._customerIdentityFromPayload(payload);
    const member = identity.member_id ? this.getMemberById(identity.member_id) : null;
    const invoices = this.outstandingInvoices
      .filter((invoice) => invoice.customer_key === identity.customer_key && toNonNegativeInt(invoice.balance_cents, 0) > 0)
      .sort((a, b) => String(a.created_at || "").localeCompare(String(b.created_at || "")) || Number(a.id) - Number(b.id));
    const lastPayment = this.customerLedger
      .filter((entry) => entry.customer_key === identity.customer_key && toNonNegativeInt(entry.debit_cents, 0) > 0)
      .slice()
      .sort((a, b) => String(b.created_at || "").localeCompare(String(a.created_at || "")) || Number(b.id) - Number(a.id))[0] || null;
    const totalCreditCents = this.customerLedger
      .filter((entry) => entry.customer_key === identity.customer_key)
      .reduce((sum, entry) => sum + toNonNegativeInt(entry.credit_cents, 0), 0);
    const totalDebitCents = this.customerLedger
      .filter((entry) => entry.customer_key === identity.customer_key)
      .reduce((sum, entry) => sum + toNonNegativeInt(entry.debit_cents, 0), 0);

    return {
      ...identity,
      member_no: member?.member_no || "",
      status: invoices.length ? "unpaid" : "paid",
      outstanding_invoice_count: invoices.length,
      last_payment_at: lastPayment?.created_at || "",
      total_credit_cents: totalCreditCents,
      total_debit_cents: totalDebitCents,
      outstanding_cents: Math.max(0, totalCreditCents - totalDebitCents),
      invoices: invoices.map((invoice) => {
        const order = this.orders.find((entry) => Number(entry.id) === Number(invoice.order_id));
        return {
          ...invoice,
          invoice_number: String(order?.invoice_number || invoice.invoice_number || invoice.receipt_number || ""),
          receipt_number: invoice.receipt_number || order?.order_number || "",
          invoice_date: invoice.created_at || order?.created_at || "",
          original_cents: toNonNegativeInt(invoice.original_cents, order?.payable_cents || order?.total_cents || 0),
          paid_cents: toNonNegativeInt(invoice.paid_cents, 0),
          balance_cents: toNonNegativeInt(invoice.balance_cents, 0),
          status: normalizeOrderStatus(invoice.status || order?.status || "unpaid", "unpaid"),
          order
        };
      })
    };
  }

  getCustomerOutstandingReport({ search = "" } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    const rowsByCustomer = new Map();

    for (const entry of this.customerLedger) {
      const key = entry.customer_key || "unknown";
      const row = rowsByCustomer.get(key) || {
        customer_key: key,
        customer_name: entry.customer_name || "Walk-in",
        phone: entry.phone || "",
        member_id: toNonNegativeInt(entry.member_id, 0),
        member_no: "",
        outstanding_invoice_count: 0,
        last_payment_at: "",
        total_credit_cents: 0,
        total_debit_cents: 0,
        current_outstanding_cents: 0
      };
      row.customer_name = row.customer_name || entry.customer_name || "Walk-in";
      row.phone = row.phone || entry.phone || "";
      row.member_id = row.member_id || toNonNegativeInt(entry.member_id, 0);
      const member = row.member_id ? this.getMemberById(row.member_id) : null;
      row.member_no = member?.member_no || row.member_no || "";
      row.total_credit_cents += toNonNegativeInt(entry.credit_cents, 0);
      row.total_debit_cents += toNonNegativeInt(entry.debit_cents, 0);
      if (toNonNegativeInt(entry.debit_cents, 0) > 0 && (!row.last_payment_at || String(entry.created_at || "") > String(row.last_payment_at))) {
        row.last_payment_at = String(entry.created_at || "");
      }
      row.current_outstanding_cents = Math.max(0, row.total_credit_cents - row.total_debit_cents);
      rowsByCustomer.set(key, row);
    }

    if (normalized) {
      for (const member of this.members.filter((entry) => entry.is_active)) {
        const matches = [member.member_no, member.name, member.phone, member.email, member.id].some((field) =>
          String(field || "").toLowerCase().includes(normalized)
        );
        if (!matches) continue;

        const identity = this._customerIdentityFromPayload({ member_id: member.id });
        if (rowsByCustomer.has(identity.customer_key)) continue;

        rowsByCustomer.set(identity.customer_key, {
          customer_key: identity.customer_key,
          customer_name: identity.customer_name || member.name || "Walk-in",
          phone: identity.phone || member.phone || "",
          member_id: member.id,
          member_no: member.member_no || "",
          outstanding_invoice_count: 0,
          last_payment_at: "",
          total_credit_cents: 0,
          total_debit_cents: 0,
          current_outstanding_cents: 0
        });
      }
    }

    for (const row of rowsByCustomer.values()) {
      row.outstanding_invoice_count = this.outstandingInvoices.filter(
        (invoice) => invoice.customer_key === row.customer_key && toNonNegativeInt(invoice.balance_cents, 0) > 0
      ).length;
      row.status = row.current_outstanding_cents > 0 ? "unpaid" : "paid";
    }

    return Array.from(rowsByCustomer.values())
      .filter((row) => {
        if (!normalized) return true;
        return [row.customer_name, row.phone, row.customer_key, row.member_no, row.member_id].some((field) =>
          String(field || "").toLowerCase().includes(normalized)
        );
      })
      .sort((a, b) => Number(b.current_outstanding_cents) - Number(a.current_outstanding_cents) || String(a.customer_name).localeCompare(String(b.customer_name)));
  }

  getCustomerLedgerReport({ search = "", customer_key = "" } = {}) {
    const normalized = String(search || "").trim().toLowerCase();
    const keyFilter = String(customer_key || "").trim();
    const runningByCustomer = new Map();
    return this._ledgerRowsForCustomer(keyFilter)
      .map((entry) => {
        const previous = runningByCustomer.get(entry.customer_key) || 0;
        const running = previous + toNonNegativeInt(entry.credit_cents, 0) - toNonNegativeInt(entry.debit_cents, 0);
        runningByCustomer.set(entry.customer_key, running);
        return {
          ...entry,
          running_balance_cents: running
        };
      })
      .filter((row) => {
        if (!normalized) return true;
        return [row.customer_name, row.phone, row.receipt_number, row.description, row.customer_key].some((field) =>
          String(field || "").toLowerCase().includes(normalized)
        );
      });
  }

  getCustomerLedgerSummary({ search = "" } = {}) {
    const outstanding = this.getCustomerOutstandingReport({ search });
    const ledger = this.getCustomerLedgerReport({ search });
    return {
      total_credit_cents: ledger.reduce((sum, row) => sum + toNonNegativeInt(row.credit_cents, 0), 0),
      total_debit_cents: ledger.reduce((sum, row) => sum + toNonNegativeInt(row.debit_cents, 0), 0),
      total_outstanding_cents: outstanding.reduce((sum, row) => sum + toNonNegativeInt(row.current_outstanding_cents, 0), 0)
    };
  }

  _normalizePendingItems(items = []) {
    if (!Array.isArray(items)) {
      return [];
    }

    return items
      .map((item) => {
        const type = String(item.type || "").trim().toLowerCase();
        const isOutstandingPayment = type === "outstanding_payment";
        if (isOutstandingPayment) {
          const amountCents = toNonNegativeInt(item.amount_cents || item.unit_price_cents, centsFromValue(item.amount || item.unit_price));
          return {
            type: "outstanding_payment",
            amount_cents: amountCents,
            customer_key: String(item.customer_key || "").trim(),
            customer_name: String(item.customer_name || "").trim(),
            phone: String(item.phone || "").trim(),
            member_id: toNonNegativeInt(item.member_id, 0)
          };
        }

        return {
          product_id: toNonNegativeInt(item.product_id || item.id, 0),
          quantity: Math.max(1, toNonNegativeInt(item.quantity, 1))
        };
      })
      .filter((item) => item.type === "outstanding_payment" ? item.amount_cents > 0 : item.product_id > 0);
  }

  _mapPendingOrder(row) {
    let items = [];
    try {
      items = JSON.parse(row.items_json || "[]");
    } catch (_) {
      items = [];
    }

    return {
      id: row.id,
      label: row.label || `Order ${row.id}`,
      customer_name: row.customer_name || "",
      phone: row.phone || "",
      member_id: toNonNegativeInt(row.member_id, 0),
      notes: row.notes || "",
      discount_id: toNonNegativeInt(row.discount_id, 0),
      items: this._normalizePendingItems(items),
      status: row.status || "active",
      created_by: row.created_by || "system",
      created_at: row.created_at,
      updated_at: row.updated_at
    };
  }

  getPendingOrders({ includeCancelled = false, actor = "" } = {}) {
    const normalizedActor = String(actor || "").trim().toLowerCase();
    return this.pendingOrders
      .filter((order) => includeCancelled || order.status !== "cancelled")
      .filter((order) => !normalizedActor || String(order.created_by || "").trim().toLowerCase() === normalizedActor)
      .slice()
      .sort((a, b) => Number(a.id) - Number(b.id))
      .map((order) => this._mapPendingOrder(order));
  }

  savePendingOrder(payload = {}, actor = "system") {
    const id = toNonNegativeInt(payload.id, 0);
    const timestamp = nowISO();
    const items = this._normalizePendingItems(payload.items);
    const base = {
      label: String(payload.label || "").trim(),
      customer_name: String(payload.customer_name || "").trim(),
      phone: String(payload.phone || "").trim(),
      member_id: toNonNegativeInt(payload.member_id, 0),
      notes: String(payload.notes || "").trim(),
      discount_id: toNonNegativeInt(payload.discount_id, 0),
      items_json: JSON.stringify(items),
      status: "active",
      updated_at: timestamp
    };

    if (id > 0) {
      const idx = this.pendingOrders.findIndex((order) => Number(order.id) === id);
      if (idx >= 0) {
        this.pendingOrders[idx] = {
          ...this.pendingOrders[idx],
          ...base,
          label: base.label || this.pendingOrders[idx].label || `Order ${id}`
        };
        this._save();
        return this._mapPendingOrder(this.pendingOrders[idx]);
      }
    }

    const nextId = this._nextId("pending_order_id");
    const order = {
      id: nextId,
      ...base,
      label: base.label || `Order ${nextId}`,
      created_by: String(actor || "system"),
      created_at: timestamp
    };
    this.pendingOrders.push(order);
    this._logAudit({
      action: "order_created",
      module: "Sales",
      changedBy: actor,
      description: `Created pending order ${order.label || nextId}`,
      details: { pending_order_id: nextId, new_status: "active" }
    });
    this._save();
    return this._mapPendingOrder(order);
  }

  cancelPendingOrder(pendingOrderId, actor = "system") {
    const id = Number(pendingOrderId);
    const idx = this.pendingOrders.findIndex((order) => Number(order.id) === id);
    if (idx < 0) {
      const err = new Error("Pending order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    this.pendingOrders[idx].status = "cancelled";
    this.pendingOrders[idx].updated_at = nowISO();
    this._logAudit({
      action: "order_cancelled",
      module: "Sales",
      changedBy: actor,
      description: `Cancelled pending order ${this.pendingOrders[idx].label || id}`,
      details: { pending_order_id: id, previous_status: "active", new_status: "cancelled" }
    });
    this._save();
    return this._mapPendingOrder(this.pendingOrders[idx]);
  }

  createOrder(payload, taxRate = 0) {
    const items = Array.isArray(payload.items) ? payload.items : [];
    if (items.length === 0) {
      const err = new Error("Order must include at least one item.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const normalizedItems = [];
    const retailItems = [];
    const outstandingPaymentItems = [];
    for (const item of items) {
      const type = String(item.type || "").trim().toLowerCase();
      const isOutstandingPayment = type === "outstanding_payment";
      if (isOutstandingPayment) {
        const amountCents = toNonNegativeInt(
          item.amount_cents || item.line_total_cents || item.unit_price_cents,
          centsFromValue(item.amount || item.line_total || item.unit_price)
        );
        const identity = this._customerIdentityFromPayload({
          customer_key: item.customer_key,
          customer_name: item.customer_name || payload.customer_name,
          phone: item.phone || payload.phone,
          member_id: item.member_id || payload.member_id
        });
        if (amountCents <= 0 || !identity.customer_key || identity.customer_key === "walk-in:unknown") {
          const err = new Error("Outstanding payment requires a selected customer and amount.");
          err.code = "BAD_REQUEST";
          throw err;
        }

        const outstandingItem = {
          product_id: 0,
          product_name: "Outstanding Payment",
          quantity: 1,
          unit_price_cents: amountCents,
          unit_cost_cents: 0,
          line_total_cents: amountCents,
          is_outstanding_payment: true,
          target_order_id: toNonNegativeInt(item.target_order_id || item.order_id, 0),
          identity
        };
        normalizedItems.push(outstandingItem);
        outstandingPaymentItems.push(outstandingItem);
        continue;
      }

      const productId = toInt(item.product_id, -1);
      const quantity = toInt(item.quantity, 0);

      if (productId <= 0 || quantity <= 0) {
        const err = new Error("Invalid product or quantity in order items.");
        err.code = "BAD_REQUEST";
        throw err;
      }

      const product = this.products.find((entry) => entry.id === productId && entry.is_active);
      if (!product) {
        const err = new Error(`Product not found: ${productId}`);
        err.code = "NOT_FOUND";
        throw err;
      }

      if (product.quantity_on_hand < quantity) {
        const err = new Error(`Insufficient stock for ${product.name}.`);
        err.code = "BAD_REQUEST";
        throw err;
      }

      const normalizedItem = {
        product_id: product.id,
        product_name: product.name,
        quantity,
        unit_price_cents: product.unit_price_cents,
        unit_cost_cents: toNonNegativeInt(product.cost_cents, 0),
        line_total_cents: product.unit_price_cents * quantity
      };
      normalizedItems.push(normalizedItem);
      retailItems.push(normalizedItem);
    }

    if (normalizedItems.length === 0) {
      const err = new Error("Order must include at least one valid item.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const rawPaymentMethod = String(payload.payment_method || "").trim().toLowerCase();
    const paymentMethod = normalizePaymentMethod(rawPaymentMethod || "cash", "cash");
    if (outstandingPaymentItems.length > 0 && paymentMethod === "havent_paid") {
      const err = new Error("Outstanding payment collection must use a real payment method.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const retailSubtotalCents = retailItems.reduce((sum, item) => sum + item.line_total_cents, 0);
    const outstandingPaymentCents = outstandingPaymentItems.reduce((sum, item) => sum + item.line_total_cents, 0);
    const subtotalCents = retailSubtotalCents + outstandingPaymentCents;
    const memberId = toNonNegativeInt(payload.member_id, 0);
    const member = memberId ? this.getMemberById(memberId) : null;
    const discountId = toNonNegativeInt(payload.discount_id || payload.discount?.id, 0);
    const selectedDiscount = discountId ? this.getDiscountById(discountId) : null;
    const discountPayload = payload.discount && typeof payload.discount === "object" ? payload.discount : null;
    let discountCents = 0;
    let discountLabel = "";
    let appliedDiscount = null;
    if (selectedDiscount) {
      const calculated = this._calculateDiscountRule(selectedDiscount, retailItems, retailSubtotalCents, member);
      discountCents = calculated.discount_cents;
      discountLabel = calculated.discount_label;
      appliedDiscount = calculated.discount;
    } else if (discountPayload) {
      const discountType = String(discountPayload.type || "").trim().toLowerCase();
      const discountValue = Number(discountPayload.value || 0);
      if (discountType === "percentage") {
        discountCents = Math.round(retailSubtotalCents * Math.max(0, discountValue) / 100);
        discountLabel = discountPayload.label || `${discountValue}% Discount`;
      } else if (discountType === "fixed") {
        discountCents = centsFromValue(discountValue);
        discountLabel = discountPayload.label || `RM${discountValue} Discount`;
      } else if (discountType === "member" && member) {
        discountCents = Math.round(retailSubtotalCents * toNonNegativeInt(member.discount_percent, 0) / 100);
        discountLabel = `${member.tier} Member ${member.discount_percent}%`;
        appliedDiscount = {
          id: 0,
          name: discountLabel,
          code: "",
          type: "member"
        };
      }
    } else if (member && payload.apply_member_discount) {
      discountCents = Math.round(retailSubtotalCents * toNonNegativeInt(member.discount_percent, 0) / 100);
      discountLabel = `${member.tier} Member ${member.discount_percent}%`;
      appliedDiscount = {
        id: 0,
        name: discountLabel,
        code: "",
        type: "member"
      };
    }
    discountCents = Math.min(retailSubtotalCents, Math.max(0, toNonNegativeInt(discountCents, 0)));
    const taxCents = Math.round(retailSubtotalCents * Number(taxRate || 0));
    const totalCents = retailSubtotalCents - discountCents + taxCents + outstandingPaymentCents;
    const payableCents = paymentMethod === "cash" && retailItems.length > 0 ? roundToNearest5Cents(totalCents) : totalCents;
    const roundingCents = payableCents - totalCents;
    const amountPaidCents = paymentMethod === "havent_paid" ? 0 : centsFromValue(payload.amount_paid ?? payableCents / 100);
    const requestedStatus = Object.hasOwn(payload, "status") ? normalizeOrderStatus(payload.status, "") : "";
    const paymentStatus =
      paymentMethod === "havent_paid"
        ? "unpaid"
        : requestedStatus ||
          (amountPaidCents <= 0 ? "unpaid" : amountPaidCents < payableCents ? "partially_paid" : "paid");
    if (paymentStatus === "paid" && amountPaidCents < payableCents) {
      const err = new Error("Paid receipts require amount paid to cover the total.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const orderId = this._nextId("order_id");
    const createdAt = nowISO();
    const orderSequence = this._nextOrderSequenceForDate(createdAt);
    const order = {
      id: orderId,
      order_number: orderNumber(orderSequence, createdAt),
      invoice_number: "",
      invoice_generated_at: "",
      invoice_pdf_file_name: "",
      invoice_pdf_file_path: "",
      invoice_pdf_generated_at: "",
      invoice_print_count: 0,
      invoice_export_count: 0,
      invoice_last_printed_at: "",
      invoice_last_exported_at: "",
      user_id: toNonNegativeInt(payload.user_id, 2),
      member_id: member ? member.id : 0,
      customer_name: String(payload.customer_name || member?.name || "").trim(),
      phone: String(payload.phone || member?.phone || "").trim(),
      status: paymentStatus,
      subtotal_cents: subtotalCents,
      tax_cents: taxCents,
      discount_cents: discountCents,
      total_cents: totalCents,
      rounding_cents: roundingCents,
      payable_cents: payableCents,
      discount_label: discountLabel,
      notes: String(payload.notes || ""),
      created_at: createdAt
    };

    for (const item of normalizedItems) {
      if (item.is_outstanding_payment) {
        this.orderItems.push({
          id: this._nextId("order_item_id"),
          order_id: orderId,
          product_id: 0,
          product_name: item.product_name,
          quantity: item.quantity,
          unit_price_cents: item.unit_price_cents,
          unit_cost_cents: 0,
          line_total_cents: item.line_total_cents
        });
        continue;
      }

      const product = this.products.find((entry) => entry.id === item.product_id);
      const beforeQty = toNonNegativeInt(product.quantity_on_hand, 0);
      product.quantity_on_hand -= item.quantity;
      product.updated_at = nowISO();
      this._logInventoryRecord({
        productId: product.id,
        action: "sale",
        quantityDelta: -item.quantity,
        beforeQty,
        afterQty: product.quantity_on_hand,
        reason: order.order_number,
        changedBy: String(payload.cashier || payload.user_id || "system")
      });

      this.orderItems.push({
        id: this._nextId("order_item_id"),
        order_id: orderId,
        product_id: item.product_id,
        product_name: item.product_name,
        quantity: item.quantity,
        unit_price_cents: item.unit_price_cents,
        unit_cost_cents: item.unit_cost_cents,
        line_total_cents: item.line_total_cents
      });
    }

    this.orders.push(order);
    this.payments.push({
      id: this._nextId("payment_id"),
      order_id: orderId,
      payment_method: paymentMethod,
      amount_cents: amountPaidCents,
      notes: String(payload.payment_notes || payload.notes || "").trim(),
      paid_at: nowISO()
    });

    const paidTowardOutstandingCents = Math.min(amountPaidCents, outstandingPaymentCents);
    let remainingOutstandingPaymentCents = paidTowardOutstandingCents;
    for (const item of outstandingPaymentItems) {
      if (remainingOutstandingPaymentCents <= 0) {
        break;
      }
      const appliedCents = Math.min(item.line_total_cents, remainingOutstandingPaymentCents);
      this._applyCustomerPayment({
        identity: item.identity,
        amountCents: appliedCents,
        orderId,
        targetOrderId: item.target_order_id,
        receiptNumber: order.order_number,
        description: `Outstanding Payment via ${order.order_number}`,
        actor: String(payload.cashier || payload.user_id || "system"),
        createdAt
      });
      remainingOutstandingPaymentCents -= appliedCents;
    }

    const paidTowardRetailCents = Math.max(0, amountPaidCents - paidTowardOutstandingCents);
    const retailPayableCents = Math.max(0, payableCents - outstandingPaymentCents);
    const retailOutstandingCents = retailItems.length > 0 ? Math.max(0, retailPayableCents - paidTowardRetailCents) : 0;
    if (retailOutstandingCents > 0) {
      this._upsertOutstandingInvoiceForOrder(order, retailOutstandingCents, String(payload.cashier || payload.user_id || "system"));
    }

    if (discountCents > 0) {
      const user = this.getUserById(order.user_id);
      this.discountUsage.push({
        id: this._nextId("discount_usage_id"),
        discount_id: toNonNegativeInt(appliedDiscount?.id, 0),
        discount_name: String(appliedDiscount?.name || discountLabel || "Discount"),
        discount_code: String(appliedDiscount?.code || ""),
        order_id: orderId,
        order_number: order.order_number,
        user_id: order.user_id,
        staff_name: user ? user.full_name || user.username : String(payload.cashier || payload.user_id || "system"),
        member_id: member ? member.id : 0,
        subtotal_cents: retailSubtotalCents,
        discount_cents: discountCents,
        created_at: createdAt
      });
    }

    this._logAudit({
      action: "order_created",
      module: "Receipt Management",
      changedBy: String(payload.cashier || payload.user_id || "system"),
      description: `Created receipt ${order.order_number}`,
      details: { order }
    });

    if (paymentStatus === "paid") {
      this._logAudit({
        action: "order_completed",
        module: "Receipt Management",
        changedBy: String(payload.cashier || payload.user_id || "system"),
        description: `Completed receipt ${order.order_number}`,
        details: { order_id: orderId, new_status: paymentStatus }
      });
    } else if (paymentStatus === "unpaid") {
      this._logAudit({
        action: "order_marked_unpaid",
        module: "Receipt Management",
        changedBy: String(payload.cashier || payload.user_id || "system"),
        description: `Marked receipt ${order.order_number} unpaid`,
        details: { order_id: orderId, new_status: paymentStatus }
      });
    }

    const pendingOrderId = toNonNegativeInt(payload.pending_order_id, 0);
    if (pendingOrderId > 0) {
      this.pendingOrders = this.pendingOrders.filter((pendingOrder) => Number(pendingOrder.id) !== pendingOrderId);
    }

    this._save();

    return {
      ...order,
      paid_cents: amountPaidCents,
      change_cents: amountPaidCents - payableCents,
      items: normalizedItems
    };
  }

  updateOrder(orderId, payload = {}, actor = "system") {
    const id = Number(orderId);
    const orderIdx = this.orders.findIndex((entry) => Number(entry.id) === id);
    if (orderIdx < 0) {
      const err = new Error("Order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const allowedStatuses = ORDER_STATUSES;
    const current = this.orders[orderIdx];
    const currentItems = this.orderItems.filter((item) => Number(item.order_id) === id).map((item) => ({ ...item }));
    const payment = this.payments
      .filter((entry) => Number(entry.order_id) === id)
      .slice()
      .sort((a, b) => Number(a.id) - Number(b.id))[0];

    let nextStatus = String(current.status || "paid").trim().toLowerCase();
    if (Object.hasOwn(payload, "status")) {
      const status = normalizeOrderStatus(payload.status, "");
      if (!allowedStatuses.has(status)) {
        const err = new Error("Invalid order status.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      nextStatus = status;
    }

    const nextNotes = Object.hasOwn(payload, "notes") ? String(payload.notes || "").trim() : String(current.notes || "");
    const nextCustomerName = Object.hasOwn(payload, "customer_name")
      ? String(payload.customer_name || "").trim()
      : String(current.customer_name || "");
    const nextPhone = Object.hasOwn(payload, "phone") ? String(payload.phone || "").trim() : String(current.phone || "");
    let nextUserId = toNonNegativeInt(current.user_id, 2);
    if (Object.hasOwn(payload, "user_id")) {
      const userId = Number(payload.user_id);
      const user = this.getUserById(userId);
      if (!user) {
        const err = new Error("Invalid staff/admin user selected.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      nextUserId = user.id;
    }

    let nextPaymentMethod = normalizePaymentMethod(payment ? payment.payment_method : "cash", "cash");
    if (Object.hasOwn(payload, "payment_method")) {
      const method = normalizePaymentMethod(payload.payment_method, "");
      if (!PAYMENT_METHODS.has(method)) {
        const err = new Error("Invalid payment method.");
        err.code = "BAD_REQUEST";
        throw err;
      }
      nextPaymentMethod = method;
    }

    const currentNormalizedItems = currentItems.map((item) => {
      const product = this.products.find((entry) => Number(entry.id) === Number(item.product_id));
      const fallbackCostCents = toNonNegativeInt(product?.cost_cents, 0);
      const quantity = toNonNegativeInt(item.quantity, 0);
      const unitPriceCents = toNonNegativeInt(item.unit_price_cents, 0);
      return {
        product_id: toNonNegativeInt(item.product_id, 0),
        product_name: String(item.product_name || ""),
        quantity,
        unit_price_cents: unitPriceCents,
        unit_cost_cents: toNonNegativeInt(item.unit_cost_cents, fallbackCostCents),
        line_total_cents: quantity * unitPriceCents
      };
    });

    let nextNormalizedItems = currentNormalizedItems;
    const hasItemsUpdate = Object.hasOwn(payload, "items");
    if (hasItemsUpdate) {
      if (!Array.isArray(payload.items) || payload.items.length === 0) {
        const err = new Error("Order must include at least one item.");
        err.code = "BAD_REQUEST";
        throw err;
      }

      nextNormalizedItems = payload.items.map((item) => {
        const productId = toInt(item.product_id, -1);
        const quantity = toInt(item.quantity, 0);
        if (productId <= 0 || quantity <= 0) {
          const err = new Error("Invalid product or quantity in order items.");
          err.code = "BAD_REQUEST";
          throw err;
        }

        const product = this.products.find((entry) => Number(entry.id) === productId);
        if (!product) {
          const err = new Error(`Product not found: ${productId}`);
          err.code = "NOT_FOUND";
          throw err;
        }

        const unitPriceCents = Object.hasOwn(item, "unit_price")
          ? centsFromValue(item.unit_price)
          : toNonNegativeInt(product.unit_price_cents, 0);

        return {
          product_id: product.id,
          product_name: String(product.name || ""),
          quantity,
          unit_price_cents: unitPriceCents,
          unit_cost_cents: toNonNegativeInt(product.cost_cents, 0),
          line_total_cents: unitPriceCents * quantity
        };
      });
    }

    if (nextNormalizedItems.length === 0) {
      const err = new Error("Order must include at least one item.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const oldQtyByProduct = new Map();
    for (const item of currentNormalizedItems) {
      const key = Number(item.product_id);
      oldQtyByProduct.set(key, (oldQtyByProduct.get(key) || 0) + toNonNegativeInt(item.quantity, 0));
    }

    const nextQtyByProduct = new Map();
    for (const item of nextNormalizedItems) {
      const key = Number(item.product_id);
      nextQtyByProduct.set(key, (nextQtyByProduct.get(key) || 0) + toNonNegativeInt(item.quantity, 0));
    }

    const stockAdjustments = [];
    const productIds = new Set([...oldQtyByProduct.keys(), ...nextQtyByProduct.keys()]);
    for (const productId of productIds) {
      const oldQty = oldQtyByProduct.get(productId) || 0;
      const newQty = nextQtyByProduct.get(productId) || 0;
      const delta = oldQty - newQty;
      if (!delta) {
        continue;
      }

      const product = this.products.find((entry) => Number(entry.id) === Number(productId));
      if (!product) {
        const err = new Error(`Product not found: ${productId}`);
        err.code = "NOT_FOUND";
        throw err;
      }

      const nextQtyOnHand = toInt(product.quantity_on_hand, 0) + delta;
      if (nextQtyOnHand < 0) {
        const err = new Error(`Insufficient stock for ${product.name}.`);
        err.code = "BAD_REQUEST";
        throw err;
      }

      stockAdjustments.push({ product, nextQtyOnHand });
    }

    const subtotalCents = nextNormalizedItems.reduce((sum, item) => sum + item.line_total_cents, 0);
    const discountCents = toNonNegativeInt(current.discount_cents, 0);
    const taxCents = 0;
    const totalCents = Math.max(0, subtotalCents - discountCents + taxCents);
    const payableCents = nextPaymentMethod === "cash" ? roundToNearest5Cents(totalCents) : totalCents;
    const roundingCents = payableCents - totalCents;

    let amountPaidCents = payment
      ? toNonNegativeInt(payment.amount_cents, toNonNegativeInt(current.payable_cents, current.total_cents))
      : toNonNegativeInt(current.payable_cents, current.total_cents);
    if (Object.hasOwn(payload, "amount_paid")) {
      amountPaidCents = centsFromValue(payload.amount_paid);
    }
    if (nextPaymentMethod === "havent_paid") {
      amountPaidCents = 0;
      nextStatus = "unpaid";
    }

    if (!Object.hasOwn(payload, "status") && Object.hasOwn(payload, "amount_paid")) {
      nextStatus = amountPaidCents <= 0 ? "unpaid" : amountPaidCents < payableCents ? "partially_paid" : "paid";
    }
    if (nextStatus === "paid" && amountPaidCents < payableCents) {
      const err = new Error("Paid receipts require amount paid to cover the total.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const beforePayableCents = toNonNegativeInt(current.payable_cents, current.total_cents);
    const beforePaidCents = payment ? toNonNegativeInt(payment.amount_cents, beforePayableCents) : beforePayableCents;
    const before = {
      ...current,
      items: currentNormalizedItems.map((item) => ({ ...item })),
      payment_method: payment ? String(payment.payment_method || "") : "",
      paid_cents: beforePaidCents,
      change_cents: beforePaidCents - beforePayableCents
    };

    for (const entry of stockAdjustments) {
      entry.product.quantity_on_hand = entry.nextQtyOnHand;
      entry.product.updated_at = nowISO();
    }

    if (hasItemsUpdate) {
      this.orderItems = this.orderItems.filter((item) => Number(item.order_id) !== id);
      for (const item of nextNormalizedItems) {
        this.orderItems.push({
          id: this._nextId("order_item_id"),
          order_id: id,
          product_id: item.product_id,
          product_name: item.product_name,
          quantity: item.quantity,
          unit_price_cents: item.unit_price_cents,
          unit_cost_cents: item.unit_cost_cents,
          line_total_cents: item.line_total_cents
        });
      }
    }

    current.status = nextStatus;
    current.notes = nextNotes;
    current.customer_name = nextCustomerName;
    current.phone = nextPhone;
    current.user_id = nextUserId;
    current.subtotal_cents = subtotalCents;
    current.tax_cents = taxCents;
    current.total_cents = totalCents;
    current.rounding_cents = roundingCents;
    current.payable_cents = payableCents;
    this.orders[orderIdx] = current;

    if (payment) {
      payment.payment_method = nextPaymentMethod;
      payment.amount_cents = amountPaidCents;
      payment.notes = String(payload.payment_notes || payload.notes || payment.notes || "").trim();
      payment.paid_at = nowISO();
    } else {
      this.payments.push({
        id: this._nextId("payment_id"),
        order_id: id,
        payment_method: nextPaymentMethod,
        amount_cents: amountPaidCents,
        notes: String(payload.payment_notes || payload.notes || "").trim(),
        paid_at: nowISO()
      });
    }

    const updatedOutstandingCents = Math.max(0, payableCents - amountPaidCents);
    const existingInvoice = this._orderOutstandingInvoice(id);
    if (updatedOutstandingCents > 0 && ["unpaid", "partially_paid"].includes(normalizeOrderStatus(nextStatus, ""))) {
      this._upsertOutstandingInvoiceForOrder(current, updatedOutstandingCents, actor);
    } else if (existingInvoice && updatedOutstandingCents <= 0) {
      existingInvoice.paid_cents = Math.max(toNonNegativeInt(existingInvoice.paid_cents, 0), toNonNegativeInt(existingInvoice.original_cents, 0));
      existingInvoice.balance_cents = 0;
      existingInvoice.status = normalizeOrderStatus(nextStatus, "paid");
      existingInvoice.updated_at = nowISO();
    }

    const after = this.getOrderById(id);
    let auditAction = before.status !== after.status ? "payment_status_changed" : "receipt_edit";
    if (after.status === "paid" && before.status !== "paid") {
      auditAction = "order_completed";
    } else if (after.status === "cancelled" && before.status !== "cancelled") {
      auditAction = "order_cancelled";
    } else if (after.status === "unpaid" && before.status !== "unpaid") {
      auditAction = "order_marked_unpaid";
    }
    this.editLogs.push({
      id: this._nextId("edit_log_id"),
      product_id: 0,
      action: auditAction,
      changed_by: String(actor || "system"),
      details: JSON.stringify({
        module: "Receipt Management",
        description:
          before.status !== after.status
            ? `Changed receipt ${current.order_number} status from ${before.status} to ${after.status}`
            : `Updated receipt ${current.order_number}`,
        order_id: id,
        previous_status: before.status,
        new_status: after.status,
        before,
        after
      }),
      created_at: nowISO()
    });

    this._save();
    return after;
  }

  collectOutstandingPayment(orderId, payload = {}, actor = "system") {
    const id = Number(orderId);
    const orderIdx = this.orders.findIndex((entry) => Number(entry.id) === id);
    if (orderIdx < 0) {
      const err = new Error("Order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const order = this.orders[orderIdx];
    const before = this.getOrderById(id);
    const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
    const paidBeforeCents = toNonNegativeInt(before?.paid_cents, 0);
    const amountCents = centsFromValue(payload.amount_paid);
    const method = normalizePaymentMethod(payload.payment_method || "cash", "cash");

    if (method === "havent_paid") {
      const err = new Error("Select a real payment method when collecting payment.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    if (amountCents <= 0) {
      const err = new Error("Payment amount must be greater than zero.");
      err.code = "BAD_REQUEST";
      throw err;
    }
    if (paidBeforeCents >= payableCents) {
      const err = new Error("This receipt is already fully paid.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const nextPaidCents = Math.min(payableCents, paidBeforeCents + amountCents);
    const nextStatus = nextPaidCents >= payableCents ? "paid" : "partially_paid";
    const previousStatus = order.status;

    this.payments.push({
      id: this._nextId("payment_id"),
      order_id: id,
      payment_method: method,
      amount_cents: Math.min(amountCents, payableCents - paidBeforeCents),
      notes: String(payload.payment_notes || payload.notes || "").trim(),
      paid_at: nowISO()
    });

    order.status = nextStatus;
    this.orders[orderIdx] = order;
    this._applyCustomerPayment({
      identity: this._customerIdentityFromOrder(order),
      amountCents: Math.min(amountCents, payableCents - paidBeforeCents),
      orderId: id,
      targetOrderId: id,
      receiptNumber: order.order_number,
      description: `Outstanding Payment for ${order.order_number}`,
      actor
    });
    const after = this.getOrderById(id);
    this._logAudit({
      action: "outstanding_payment_collected",
      module: "Outstanding Orders",
      changedBy: actor,
      description: `Collected outstanding payment for ${order.order_number}`,
      details: {
        order_id: id,
        previous_status: previousStatus,
        new_status: nextStatus,
        amount_cents: Math.min(amountCents, payableCents - paidBeforeCents),
        before,
        after
      }
    });
    if (previousStatus !== nextStatus) {
      this._logAudit({
        action: nextStatus === "paid" ? "order_completed" : "payment_status_changed",
        module: "Outstanding Orders",
        changedBy: actor,
        description: `Changed receipt ${order.order_number} status from ${previousStatus} to ${nextStatus}`,
        details: {
          order_id: id,
          previous_status: previousStatus,
          new_status: nextStatus
        }
      });
    }

    this._save();
    return after;
  }

  clearOutstandingInvoice(orderId, actor = "system") {
    const id = Number(orderId);
    const orderIdx = this.orders.findIndex((entry) => Number(entry.id) === id);
    if (orderIdx < 0) {
      const err = new Error("Order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const before = this.getOrderById(id);
    const order = this.orders[orderIdx];
    const invoice = this._orderOutstandingInvoice(id);
    const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
    const paidCents = toNonNegativeInt(before?.paid_cents, 0);
    const balanceCents = invoice ? toNonNegativeInt(invoice.balance_cents, 0) : Math.max(0, payableCents - paidCents);

    if (balanceCents > 0) {
      const err = new Error("Outstanding balance must be RM0.00 before clearing.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    order.status = "paid";
    order.updated_at = nowISO();
    if (invoice) {
      invoice.balance_cents = 0;
      invoice.paid_cents = Math.max(toNonNegativeInt(invoice.paid_cents, 0), toNonNegativeInt(invoice.original_cents, payableCents));
      invoice.status = "paid";
      invoice.updated_at = nowISO();
    }
    this.orders[orderIdx] = order;

    const after = this.getOrderById(id);
    this._logAudit({
      action: "outstanding_cleared",
      module: "Outstanding Orders",
      changedBy: actor,
      description: `Cleared outstanding invoice ${order.invoice_number || order.order_number || id}`,
      details: {
        order_id: id,
        invoice_number: order.invoice_number || "",
        receipt_number: order.order_number || "",
        customer_name: before?.customer_name || before?.member?.name || order.customer_name || "",
        amount_paid_cents: paidCents,
        before,
        after
      }
    });

    this._save();
    return after;
  }

  _outstandingRow(order) {
    const fullOrder = this.getOrderById(order.id);
    const member = fullOrder?.member || null;
    const user = fullOrder?.user || null;
    const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
    const paidCents = toNonNegativeInt(fullOrder?.paid_cents, 0);
    const balanceCents = Math.max(0, payableCents - paidCents);
    const createdAt = String(order.created_at || nowISO());
    const dueDate = dateKeyLocal(new Date(new Date(createdAt).getTime() + 7 * 24 * 60 * 60 * 1000));

    return {
      ...order,
      invoice_number: String(order.invoice_number || ""),
      customer_name: member?.name || order.customer_name || "Walk-in",
      phone: member?.phone || order.phone || "",
      description: order.notes || "Retail Sale",
      payment_method: fullOrder?.payment_method || "",
      paid_cents: paidCents,
      outstanding_cents: balanceCents,
      due_date: dueDate,
      cashier: user ? user.full_name || user.username : "",
      user_name: user ? user.full_name || user.username : "",
      member
    };
  }

  getOutstandingOrders({ search = "", filter = "" } = {}) {
    const normalizedSearch = String(search || "").trim().toLowerCase();
    const normalizedFilter = String(filter || "").trim().toLowerCase();
    const today = dateKeyLocal(new Date());

    return this.outstandingInvoices
      .filter((invoice) => toNonNegativeInt(invoice.balance_cents, 0) > 0)
      .map((invoice) => {
        const order = this.orders.find((entry) => Number(entry.id) === Number(invoice.order_id));
        const fullOrder = order ? this.getOrderById(order.id) : null;
        const user = order ? this.getUserById(order.user_id) : null;
        const payableCents = order ? toNonNegativeInt(order.payable_cents, order.total_cents) : toNonNegativeInt(invoice.original_cents, 0);
        const balanceCents = toNonNegativeInt(invoice.balance_cents, 0);
        return {
          ...(order || {}),
          id: order?.id || invoice.order_id || invoice.id,
          outstanding_invoice_id: invoice.id,
          order_number: invoice.receipt_number || order?.order_number || "",
          invoice_number: String(order?.invoice_number || ""),
          customer_key: invoice.customer_key,
          customer_name: invoice.customer_name || order?.customer_name || "Walk-in",
          phone: invoice.phone || order?.phone || "",
          description: invoice.description || order?.notes || "Retail Sale",
          payment_method: fullOrder?.payment_method || "",
          paid_cents: Math.max(0, payableCents - balanceCents),
          outstanding_cents: balanceCents,
          payable_cents: payableCents,
          total_cents: order ? toNonNegativeInt(order.total_cents, payableCents) : payableCents,
          status: normalizeOrderStatus(invoice.status || order?.status || "unpaid", "unpaid"),
          created_at: invoice.created_at || order?.created_at || nowISO(),
          due_date: this._invoiceDueDate(invoice),
          cashier: user ? user.full_name || user.username : "",
          user_name: user ? user.full_name || user.username : "",
          member: invoice.member_id ? this.getMemberById(invoice.member_id) : null
        };
      })
      .filter((order) => {
        if (!normalizedFilter || normalizedFilter === "all") return true;
        if (normalizedFilter === "overdue") return String(order.due_date || "") < today;
        return normalizeOrderStatus(order.status, "") === normalizedFilter;
      })
      .filter((order) => {
        if (!normalizedSearch) return true;
        return [order.customer_name, order.phone, order.order_number, order.invoice_number, order.description]
          .some((field) => String(field || "").toLowerCase().includes(normalizedSearch));
      })
      .sort((a, b) => String(a.due_date || "").localeCompare(String(b.due_date || "")) || Number(b.id) - Number(a.id));
  }

  getOutstandingSummary(limit = 5) {
    const rows = this.getOutstandingOrders();
    const today = dateKeyLocal(new Date());
    const todayRows = rows.filter((row) => dateKeyLocal(row.created_at) === today);
    const recentPayments = this.payments
      .slice()
      .sort((a, b) => Number(b.id) - Number(a.id))
      .filter((payment) => toNonNegativeInt(payment.amount_cents, 0) > 0)
      .filter((payment) => normalizePaymentMethod(payment.payment_method, "") !== "havent_paid")
      .slice(0, Math.max(1, Math.min(20, toNonNegativeInt(limit, 5))))
      .map((payment) => {
        const order = this.orders.find((entry) => Number(entry.id) === Number(payment.order_id));
        return {
          ...payment,
          order_number: order?.order_number || "",
          status: order?.status || ""
        };
      });

    return {
      outstanding_count: rows.length,
      todays_unpaid_cents: todayRows.reduce((sum, row) => sum + toNonNegativeInt(row.outstanding_cents, 0), 0),
      outstanding_cents: rows.reduce((sum, row) => sum + toNonNegativeInt(row.outstanding_cents, 0), 0),
      overdue_count: rows.filter((row) => String(row.due_date || "") < today).length,
      recent_payments: recentPayments
    };
  }

  deleteOrder(orderId, actor = "admin") {
    const id = Number(orderId);
    const before = this.getOrderById(id);
    if (!before) {
      const err = new Error("Order not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    this.orders = this.orders.filter((order) => Number(order.id) !== id);
    this.orderItems = this.orderItems.filter((item) => Number(item.order_id) !== id);
    this.payments = this.payments.filter((payment) => Number(payment.order_id) !== id);
    this.discountUsage = this.discountUsage.filter((usage) => Number(usage.order_id) !== id);
    this.outstandingInvoices = this.outstandingInvoices.filter((invoice) => Number(invoice.order_id) !== id);
    this.customerLedger = this.customerLedger.filter((entry) => Number(entry.order_id) !== id);
    this._logAudit({
      action: "receipt_delete",
      module: "Receipt Management",
      changedBy: actor,
      description: `Deleted receipt ${before.order_number || id}`,
      details: { before }
    });
    this._save();
    return before;
  }

  getOrderById(id) {
    const order = this.orders.find((entry) => entry.id === Number(id));
    if (!order) {
      return null;
    }

    const items = this.orderItems
      .filter((item) => item.order_id === order.id)
      .map((item) => ({ ...item }));

    const payment = this.getPaymentByOrderId(order.id);
    const user = this.getUserById(order.user_id);
    const member = this.getMemberById(order.member_id);
    const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
    const invoice = this._orderOutstandingInvoice(order.id);
    const invoiceBalanceCents = invoice ? toNonNegativeInt(invoice.balance_cents, 0) : null;
    const paidCents = invoice
      ? Math.max(0, payableCents - invoiceBalanceCents)
      : payment
        ? toNonNegativeInt(payment.amount_cents, payableCents)
        : payableCents;
    const changeCents = paidCents - payableCents;

    return {
      ...order,
      status: invoice ? normalizeOrderStatus(invoice.status || order.status, order.status) : order.status,
      items,
      payment_method: payment ? payment.payment_method : "",
      paid_cents: paidCents,
      change_cents: changeCents,
      user: user
        ? {
            id: user.id,
            username: user.username,
            full_name: user.full_name,
            role: user.role
          }
        : null,
      member: member
        ? {
            id: member.id,
            member_no: member.member_no,
            name: member.name,
            phone: member.phone,
            email: member.email,
            tier: member.tier,
            discount_percent: member.discount_percent
          }
        : null
    };
  }

  _getOrderRows(search = "", statusFilter = "") {
    const normalizedSearch = String(search || "").trim().toLowerCase();
    const normalizedStatus = String(statusFilter || "").trim().toLowerCase().replace(/[\s-]+/g, "_");
    const rows = this.orders
      .slice()
      .sort((a, b) => b.id - a.id)
      .map((order) => {
        const payment = this.getPaymentByOrderId(order.id);
        const user = this.getUserById(order.user_id);
        const member = this.getMemberById(order.member_id);
        const items = this.orderItems.filter((item) => Number(item.order_id) === Number(order.id));
        const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
        const invoice = this._orderOutstandingInvoice(order.id);
        const outstandingCents = invoice ? toNonNegativeInt(invoice.balance_cents, 0) : Math.max(0, payableCents - (payment ? toNonNegativeInt(payment.amount_cents, 0) : payableCents));
        const paidCents = invoice ? Math.max(0, payableCents - outstandingCents) : payment ? toNonNegativeInt(payment.amount_cents, 0) : payableCents;
        const paymentStatus = invoice ? normalizeOrderStatus(invoice.status || order.status, "paid") : normalizeOrderStatus(order.status, "paid");

        const row = {
          ...order,
          status: paymentStatus,
          invoice_number: String(order.invoice_number || ""),
          item_count: items.length,
          payment_method: payment ? payment.payment_method : "",
          paid_cents: paidCents,
          outstanding_cents: outstandingCents,
          description: order.notes || "Retail Sale",
          customer_name: member?.name || order.customer_name || "Walk-in",
          phone: member?.phone || order.phone || "",
          user_name: user ? user.full_name || user.username : "",
          user_role: user ? user.role : ""
        };

        if (normalizedStatus && normalizedStatus !== "all") {
          if (normalizedStatus === "outstanding") {
            if (outstandingCents <= 0 || !["unpaid", "partially_paid"].includes(paymentStatus)) {
              return { row, matches: false };
            }
          } else if (paymentStatus !== normalizedStatus) {
            return { row, matches: false };
          }
        }

        if (!normalizedSearch) {
          return { row, matches: true };
        }

        const hasItemMatch = items.some((item) =>
          String(item.product_name || "")
            .toLowerCase()
            .includes(normalizedSearch)
        );

        const fields = [
          row.id,
          row.order_number,
          row.invoice_number,
          row.status,
          row.payment_method,
          row.description,
          row.customer_name,
          row.phone,
          row.user_name,
          row.user_role,
          row.created_at
        ];

        const hasRowMatch = fields.some((field) => String(field || "").toLowerCase().includes(normalizedSearch));
        return { row, matches: hasRowMatch || hasItemMatch };
      });

    return rows.filter((entry) => entry.matches).map((entry) => entry.row);
  }

  getOrders(limit = 200, search = "", status = "") {
    const max = toNonNegativeInt(limit, 200) || 200;
    return this._getOrderRows(search, status).slice(0, max);
  }

  getOrdersPage({ page = 1, pageSize = 10, search = "", status = "" } = {}) {
    const safePageSize = Math.min(200, Math.max(1, toNonNegativeInt(pageSize, 10) || 10));
    const sorted = this._getOrderRows(search, status);
    const total = sorted.length;
    const totalPages = Math.max(1, Math.ceil(total / safePageSize));
    const requestedPage = Math.max(1, toNonNegativeInt(page, 1) || 1);
    const safePage = Math.min(requestedPage, totalPages);
    const start = (safePage - 1) * safePageSize;
    const rows = sorted.slice(start, start + safePageSize);

    return {
      rows,
      pagination: {
        page: safePage,
        page_size: safePageSize,
        total,
        total_pages: totalPages
      }
    };
  }

  getEditLogs(limit = 500) {
    const max = toNonNegativeInt(limit, 500) || 500;
    return this.editLogs
      .slice()
      .sort((a, b) => b.id - a.id)
      .slice(0, max)
      .map((log) => ({ ...log }));
  }

  getEditLogsPage({ page = 1, pageSize = 10 } = {}) {
    const safePageSize = Math.min(200, Math.max(1, toNonNegativeInt(pageSize, 10) || 10));
    const sorted = this.editLogs.slice().sort((a, b) => b.id - a.id);
    const total = sorted.length;
    const totalPages = Math.max(1, Math.ceil(total / safePageSize));
    const requestedPage = Math.max(1, toNonNegativeInt(page, 1) || 1);
    const safePage = Math.min(requestedPage, totalPages);
    const start = (safePage - 1) * safePageSize;

    const rows = sorted.slice(start, start + safePageSize).map((log) => ({ ...log }));
    return {
      rows,
      pagination: {
        page: safePage,
        page_size: safePageSize,
        total,
        total_pages: totalPages
      }
    };
  }

  getStock(includeInactive = true, search = "") {
    const normalized = String(search || "").trim().toLowerCase();
    return this.products
      .filter((product) => includeInactive || product.is_active)
      .filter((product) => {
        if (!normalized) {
          return true;
        }
        return [product.sku, product.name, product.barcode, product.category, product.subcategory, product.brand, product.supplier]
          .map((field) => String(field || "").toLowerCase())
          .some((value) => value.includes(normalized));
      })
      .slice()
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((product) => ({
        id: product.id,
        sku: product.sku,
        name: product.name,
        category: product.category,
        subcategory: product.subcategory,
        brand: product.brand || "",
        supplier: product.supplier || "",
        quantity_on_hand: product.quantity_on_hand,
        reorder_level: product.reorder_level,
        is_active: product.is_active,
        updated_at: product.updated_at
      }));
  }

  getLowStockProducts(limit = 20) {
    const max = Math.min(200, Math.max(1, toNonNegativeInt(limit, 20) || 20));
    return this.products
      .filter((product) => product.is_active && toNonNegativeInt(product.quantity_on_hand, 0) <= toNonNegativeInt(product.reorder_level, 0))
      .slice()
      .sort(
        (a, b) =>
          toNonNegativeInt(a.quantity_on_hand, 0) - toNonNegativeInt(b.quantity_on_hand, 0) ||
          String(a.name || "").localeCompare(String(b.name || ""))
      )
      .slice(0, max)
      .map((product) => ({ ...product }));
  }

  addStock(productId, quantityToAdd, actor = "system") {
    const id = Number(productId);
    const qty = toInt(quantityToAdd, 0);

    if (!Number.isInteger(id) || id <= 0) {
      const err = new Error("Invalid product ID.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    if (!Number.isInteger(qty) || qty <= 0) {
      const err = new Error("Quantity to add must be a positive integer.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    const idx = this.products.findIndex((product) => Number(product.id) === id);
    if (idx < 0) {
      const err = new Error("Product not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const before = { ...this.products[idx] };
    this.products[idx].quantity_on_hand = toNonNegativeInt(this.products[idx].quantity_on_hand, 0) + qty;
    this.products[idx].updated_at = nowISO();
    this._logInventoryRecord({
      productId: this.products[idx].id,
      action: "stock_add",
      quantityDelta: qty,
      beforeQty: before.quantity_on_hand,
      afterQty: this.products[idx].quantity_on_hand,
      reason: "Stock adjustment",
      changedBy: actor
    });

    this._logEdit({
      productId: this.products[idx].id,
      action: "stock_add",
      changedBy: actor,
      details: JSON.stringify({
        added_qty: qty,
        before: {
          quantity_on_hand: before.quantity_on_hand
        },
        after: {
          quantity_on_hand: this.products[idx].quantity_on_hand
        }
      })
    });

    this._save();
    return { ...this.products[idx] };
  }

  getInventoryRecords({ productId = 0, limit = 500 } = {}) {
    const id = Number(productId);
    const max = Math.min(1000, Math.max(1, toNonNegativeInt(limit, 500) || 500));
    return this.inventoryRecords
      .filter((record) => !id || Number(record.product_id) === id)
      .slice()
      .sort((a, b) => Number(b.id) - Number(a.id))
      .slice(0, max)
      .map((record) => ({ ...record }));
  }

  logAudit({ action, module, changedBy, description, details = {} }) {
    this._logAudit({ action, module, changedBy, description, details });
    this._save();
  }

  _nextId(key) {
    this.meta[key] = toNonNegativeInt(this.meta[key], 0) + 1;
    return this.meta[key];
  }

  _logAudit({ action, module, changedBy, description, details = {} }) {
    const actor = this.users.find((user) => String(user.username || "").toLowerCase() === String(changedBy || "").toLowerCase());
    this._logEdit({
      productId: 0,
      action: String(action || "audit"),
      changedBy: String(changedBy || "system"),
      details: JSON.stringify({
        module: String(module || "System"),
        description: String(description || ""),
        role: actor?.role || "",
        ...details
      })
    });
  }

  _logEdit({ productId, action, changedBy, details }) {
    this.editLogs.push({
      id: this._nextId("edit_log_id"),
      product_id: Number(productId),
      action: String(action || "edit"),
      changed_by: String(changedBy || "system"),
      details: String(details || ""),
      created_at: nowISO()
    });
  }

  _logInventoryRecord({ productId, action, quantityDelta, beforeQty, afterQty, reason, changedBy }) {
    this.inventoryRecords.push({
      id: this._nextId("inventory_record_id"),
      product_id: Number(productId),
      action: String(action || "adjust"),
      quantity_delta: toInt(quantityDelta, 0),
      before_qty: toNonNegativeInt(beforeQty, 0),
      after_qty: toNonNegativeInt(afterQty, 0),
      reason: String(reason || ""),
      changed_by: String(changedBy || "system"),
      created_at: nowISO()
    });
  }

  _seedDefaults() {
    const createdAt = nowISO();
    this.products = [
      {
        id: 1,
        sku: "DRK-COLA-330",
        name: "Cola 330ml",
        category: "Drinks",
        subcategory: "Soft Drink",
        barcode: "100000000001",
        brand: "House",
        supplier: "House Supplier",
        description: "Canned soft drink",
        template_type: "sports-drink",
        attributes_json: JSON.stringify({ volume: "330ml" }),
        unit_price_cents: 199,
        cost_cents: 90,
        quantity_on_hand: 50,
        reorder_level: 10,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 2,
        sku: "DRK-WATER-500",
        name: "Mineral Water 500ml",
        category: "Drinks",
        subcategory: "Water",
        barcode: "100000000002",
        brand: "House",
        supplier: "House Supplier",
        description: "Bottled mineral water",
        template_type: "sports-drink",
        attributes_json: JSON.stringify({ volume: "500ml" }),
        unit_price_cents: 129,
        cost_cents: 50,
        quantity_on_hand: 80,
        reorder_level: 20,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 3,
        sku: "SNK-CHIPS-SEA",
        name: "Sea Salt Chips",
        category: "Snacks",
        subcategory: "Chips",
        barcode: "100000000003",
        brand: "House",
        supplier: "House Supplier",
        description: "Sea salt chips",
        template_type: "accessories",
        attributes_json: JSON.stringify({ category: "Snack" }),
        unit_price_cents: 249,
        cost_cents: 120,
        quantity_on_hand: 40,
        reorder_level: 8,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 4,
        sku: "SNK-CHOCO-BAR",
        name: "Chocolate Bar",
        category: "Snacks",
        subcategory: "Chocolate",
        barcode: "100000000004",
        brand: "House",
        supplier: "House Supplier",
        description: "Chocolate snack bar",
        template_type: "accessories",
        attributes_json: JSON.stringify({ category: "Snack" }),
        unit_price_cents: 159,
        cost_cents: 70,
        quantity_on_hand: 60,
        reorder_level: 12,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 5,
        sku: "BAK-CROISSANT",
        name: "Butter Croissant",
        category: "Bakery",
        subcategory: "Pastry",
        barcode: "100000000005",
        brand: "House",
        supplier: "House Supplier",
        description: "Butter pastry",
        template_type: "accessories",
        attributes_json: JSON.stringify({ category: "Snack" }),
        unit_price_cents: 299,
        cost_cents: 130,
        quantity_on_hand: 25,
        reorder_level: 6,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 6,
        sku: "BAK-MUFFIN-BLU",
        name: "Blueberry Muffin",
        category: "Bakery",
        subcategory: "Muffin",
        barcode: "100000000006",
        brand: "House",
        supplier: "House Supplier",
        description: "Blueberry muffin",
        template_type: "accessories",
        attributes_json: JSON.stringify({ category: "Snack" }),
        unit_price_cents: 279,
        cost_cents: 110,
        quantity_on_hand: 30,
        reorder_level: 6,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      }
    ];

    this.orders = [];
    this.orderItems = [];
    this.payments = [];
    this.pendingOrders = [];
    this.editLogs = [];
    this.outstandingInvoices = [];
    this.customerLedger = [];
    this.settingsRows = [];
    this.suppliers = [
      {
        id: 1,
        name: "House Supplier",
        contact_person: "",
        phone: "",
        email: "",
        address: "",
        payment_terms: "COD",
        notes: "Default supplier for seeded inventory.",
        outstanding_cents: 0,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      }
    ];
    this.purchaseOrders = [];
    this.purchaseOrderItems = [];
    this.expenses = [];
    this.dailyClosings = [];
    this.supplierPayments = [];
    this.members = [
      {
        id: 1,
        member_no: "MBR-00001",
        name: "Sample Silver Member",
        phone: "0100000001",
        email: "",
        tier: "Silver",
        discount_percent: 5,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 2,
        member_no: "MBR-00002",
        name: "Sample Gold Member",
        phone: "0100000002",
        email: "",
        tier: "Gold",
        discount_percent: 10,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      },
      {
        id: 3,
        member_no: "MBR-00003",
        name: "Sample VIP Member",
        phone: "0100000003",
        email: "",
        tier: "VIP",
        discount_percent: 15,
        is_active: 1,
        created_at: createdAt,
        updated_at: createdAt
      }
    ];
    this.discounts = [
      { id: 1, name: "5%", type: "percentage", scope: "cart", value: 5, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 2, name: "10%", type: "percentage", scope: "cart", value: 10, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 3, name: "15%", type: "percentage", scope: "cart", value: 15, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 4, name: "20%", type: "percentage", scope: "cart", value: 20, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 5, name: "RM5", type: "fixed", scope: "cart", value: 5, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 6, name: "RM10", type: "fixed", scope: "cart", value: 10, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 7, name: "RM20", type: "fixed", scope: "cart", value: 20, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt }
    ];
    this.discountUsage = [];
    this.inventoryRecords = [];
    this.loginHistory = [];
    this.users = [];
    this._ensureUsersSeeded(createdAt);
    this.meta = {
      product_id: 6,
      order_id: 0,
      order_item_id: 0,
      payment_id: 0,
      pending_order_id: 0,
      edit_log_id: 0,
      member_id: 3,
      supplier_id: 1,
      purchase_order_id: 0,
      purchase_order_item_id: 0,
      expense_id: 0,
      daily_closing_id: 0,
      supplier_payment_id: 0,
      discount_id: 7,
      discount_usage_id: 0,
      inventory_record_id: 0,
      outstanding_invoice_id: 0,
      customer_ledger_id: 0,
      login_history_id: 0
    };
  }

  _defaultSystemUser(account, createdAt = nowISO()) {
    return {
      id: Math.max(...this.users.map((item) => Number(item.id || 0)), 0) + 1,
      username: account.username,
      password: hashPasswordSecure(account.password),
      full_name: account.full_name,
      role: normalizeUserRole(account.role, "admin"),
      email: "",
      phone: "",
      employee_id: "",
      profile_photo: "",
      status: "active",
      is_active: 1,
      force_password_change: 1,
      last_login_at: "",
      created_at: createdAt,
      updated_at: createdAt
    };
  }

  _ensureUsersSeeded(createdAt = nowISO()) {
    for (const account of DEFAULT_SYSTEM_USERS) {
      if (isSystemOwnerRole(account.role)) {
        const activeOwner = this.users.find((user) =>
          isSystemOwnerRole(user.role) &&
          normalizeUserStatus(user.status, user.is_active) === "active"
        );
        if (activeOwner) {
          continue;
        }
        const existingOwner = this.users.find((user) => isSystemOwnerRole(user.role));
        if (existingOwner) {
          existingOwner.status = "active";
          existingOwner.is_active = 1;
          existingOwner.updated_at = existingOwner.updated_at || createdAt;
          continue;
        }
        const usernameMatch = this.users.find((user) => String(user.username || "").toLowerCase() === account.username.toLowerCase());
        if (usernameMatch) {
          usernameMatch.username = account.username;
          usernameMatch.password = hashPasswordSecure(account.password);
          usernameMatch.full_name = usernameMatch.full_name || account.full_name;
          usernameMatch.role = "system_owner";
          usernameMatch.status = "active";
          usernameMatch.is_active = 1;
          usernameMatch.force_password_change = 1;
          usernameMatch.updated_at = createdAt;
          continue;
        }
      }
      const exists = this.users.some((user) => String(user.username || "").toLowerCase() === account.username.toLowerCase());
      if (!exists) {
        this.users.push(this._defaultSystemUser(account, createdAt));
      }
    }
  }

  _ensureDiscountsSeeded() {
    if (this.discounts.length > 0) {
      return;
    }

    const createdAt = nowISO();
    this.discounts = [
      { id: 1, name: "5%", type: "percentage", scope: "cart", value: 5, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 2, name: "10%", type: "percentage", scope: "cart", value: 10, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 3, name: "15%", type: "percentage", scope: "cart", value: 15, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 4, name: "20%", type: "percentage", scope: "cart", value: 20, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 5, name: "RM5", type: "fixed", scope: "cart", value: 5, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 6, name: "RM10", type: "fixed", scope: "cart", value: 10, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt },
      { id: 7, name: "RM20", type: "fixed", scope: "cart", value: 20, member_tier: "", promo_type: "", is_active: 1, created_by: "system", created_at: createdAt, updated_at: createdAt }
    ];
  }

  _ensureSuppliersSeeded() {
    const existing = new Set(this.suppliers.map((supplier) => String(supplier.name || "").trim().toLowerCase()).filter(Boolean));
    const names = Array.from(
      new Set(this.products.map((product) => String(product.supplier || "").trim()).filter(Boolean))
    );
    if (!names.length && this.suppliers.length > 0) {
      return;
    }
    const timestamp = nowISO();
    for (const name of names.length ? names : ["House Supplier"]) {
      if (existing.has(name.toLowerCase())) {
        continue;
      }
      const id = this._nextId("supplier_id");
      this.suppliers.push({
        id,
        name,
        contact_person: "",
        phone: "",
        email: "",
        address: "",
        payment_terms: "COD",
        notes: names.length ? "Imported from product supplier field." : "Default supplier.",
        outstanding_cents: 0,
        is_active: 1,
        created_at: timestamp,
        updated_at: timestamp
      });
      existing.add(name.toLowerCase());
    }
  }

  _ensureOutstandingLedgerSeeded() {
    for (const order of this.orders) {
      const payment = this.getPaymentByOrderId(order.id);
      const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
      const paidCents = payment ? toNonNegativeInt(payment.amount_cents, 0) : payableCents;
      const balanceCents = Math.max(0, payableCents - paidCents);
      const status = normalizeOrderStatus(order.status, "paid");
      if (balanceCents <= 0 || !["unpaid", "partially_paid"].includes(status)) {
        continue;
      }
      this._upsertOutstandingInvoiceForOrder(order, balanceCents, "migration");
    }
  }

  _ensureSettingsSeeded() {
    const timestamp = nowISO();
    const defaults = {
      storage_engine: "excel",
      workbook_path: this.filePath,
      data_folder: path.dirname(this.filePath),
      workbook_version: "2",
      invoice_settings_json: JSON.stringify({
        prefix: "Invoice",
        date_format: "YYMM",
        number_length: 3,
        reset: "monthly"
      }),
      pdf_export_settings_json: JSON.stringify({
        folder: "downloads",
        custom_path: ""
      }),
      invoice_sequence_json: "{}",
      last_invoice_number: ""
    };
    const existingKeys = new Set(this.settingsRows.map((row) => row.key));
    for (const [key, value] of Object.entries(defaults)) {
      if (!existingKeys.has(key)) {
        this.settingsRows.push({ key, value, updated_at: timestamp });
      }
    }
  }

  _ensureCompanySettingsSeeded() {
    const timestamp = nowISO();
    const existing = this.companySettings[0];
    if (existing) {
      const normalized = normalizeCompanySettings(existing, existing);
      const changed = JSON.stringify(normalized) !== JSON.stringify(existing);
      this.companySettings = [normalized];
      return changed;
    }

    let template = {};
    try {
      template = JSON.parse(this.getSetting("receipt_template_json", "{}") || "{}");
    } catch (_) {
      template = {};
    }
    const defaults = defaultCompanySettings(timestamp);
    this.companySettings = [
      normalizeCompanySettings(
        {
          ...defaults,
          company_logo: template.company_logo || defaults.company_logo,
          business_address: template.address || defaults.business_address,
          phone: template.phone || defaults.phone,
          email: template.email || defaults.email
        },
        defaults
      )
    ];
    return true;
  }

  _salesRowsForWorkbook() {
    return this.orders.map((order) => {
      const payment = this.getPaymentByOrderId(order.id);
      const user = this.getUserById(order.user_id);
      const invoice = this._orderOutstandingInvoice(order.id);
      const payableCents = toNonNegativeInt(order.payable_cents, order.total_cents);
      const outstandingCents = invoice ? toNonNegativeInt(invoice.balance_cents, 0) : Math.max(0, payableCents - (payment ? toNonNegativeInt(payment.amount_cents, 0) : payableCents));
      return {
        id: order.id,
        order_number: order.order_number,
        invoice_number: String(order.invoice_number || ""),
        customer_name: order.customer_name || "Walk-in",
        phone: order.phone || "",
        status: invoice ? normalizeOrderStatus(invoice.status || order.status, order.status) : order.status,
        subtotal_cents: order.subtotal_cents,
        discount_cents: order.discount_cents,
        tax_cents: order.tax_cents,
        payable_cents: payableCents,
        paid_cents: Math.max(0, payableCents - outstandingCents),
        outstanding_cents: outstandingCents,
        payment_method: payment ? payment.payment_method : "",
        cashier: user ? user.full_name || user.username : "",
        notes: order.notes || "",
        created_at: order.created_at
      };
    });
  }

  _rebuildMeta() {
    this.meta.product_id = Math.max(this.meta.product_id, ...this.products.map((item) => item.id), 0);
    this.meta.order_id = Math.max(this.meta.order_id, ...this.orders.map((item) => item.id), 0);
    this.meta.order_item_id = Math.max(this.meta.order_item_id, ...this.orderItems.map((item) => item.id), 0);
    this.meta.payment_id = Math.max(this.meta.payment_id, ...this.payments.map((item) => item.id), 0);
    this.meta.pending_order_id = Math.max(
      this.meta.pending_order_id || 0,
      ...this.pendingOrders.map((item) => item.id),
      0
    );
    this.meta.edit_log_id = Math.max(this.meta.edit_log_id, ...this.editLogs.map((item) => item.id), 0);
    this.meta.member_id = Math.max(this.meta.member_id || 0, ...this.members.map((item) => item.id), 0);
    this.meta.supplier_id = Math.max(this.meta.supplier_id || 0, ...this.suppliers.map((item) => item.id), 0);
    this.meta.purchase_order_id = Math.max(this.meta.purchase_order_id || 0, ...this.purchaseOrders.map((item) => item.id), 0);
    this.meta.purchase_order_item_id = Math.max(
      this.meta.purchase_order_item_id || 0,
      ...this.purchaseOrderItems.map((item) => item.id),
      0
    );
    this.meta.expense_id = Math.max(this.meta.expense_id || 0, ...this.expenses.map((item) => item.id), 0);
    this.meta.daily_closing_id = Math.max(this.meta.daily_closing_id || 0, ...this.dailyClosings.map((item) => item.id), 0);
    this.meta.supplier_payment_id = Math.max(
      this.meta.supplier_payment_id || 0,
      ...this.supplierPayments.map((item) => item.id),
      0
    );
    this.meta.discount_id = Math.max(this.meta.discount_id || 0, ...this.discounts.map((item) => item.id), 0);
    this.meta.discount_usage_id = Math.max(
      this.meta.discount_usage_id || 0,
      ...this.discountUsage.map((item) => item.id),
      0
    );
    this.meta.inventory_record_id = Math.max(
      this.meta.inventory_record_id || 0,
      ...this.inventoryRecords.map((item) => item.id),
      0
    );
    this.meta.outstanding_invoice_id = Math.max(
      this.meta.outstanding_invoice_id || 0,
      ...this.outstandingInvoices.map((item) => item.id),
      0
    );
    this.meta.customer_ledger_id = Math.max(
      this.meta.customer_ledger_id || 0,
      ...this.customerLedger.map((item) => item.id),
      0
    );
    this.meta.login_history_id = Math.max(
      this.meta.login_history_id || 0,
      ...this.loginHistory.map((item) => item.id),
      0
    );
  }

  _readSheet(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      return [];
    }

    return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  }

  _save() {
    const workbook = XLSX.utils.book_new();
    this._appendSheet(workbook, "products", this.products, HEADERS.products);
    this._appendSheet(workbook, "orders", this.orders, HEADERS.orders);
    this._appendSheet(workbook, "order_items", this.orderItems, HEADERS.order_items);
    this._appendSheet(workbook, "payments", this.payments, HEADERS.payments);
    this._appendSheet(workbook, "pending_orders", this.pendingOrders, HEADERS.pending_orders);
    this._appendSheet(workbook, "edit_logs", this.editLogs, HEADERS.edit_logs);
    this._appendSheet(workbook, "users", this.users, HEADERS.users);
    this._appendSheet(workbook, "login_history", this.loginHistory, HEADERS.login_history);
    this._appendSheet(workbook, "members", this.members, HEADERS.members);
    this._appendSheet(workbook, "suppliers", this.suppliers, HEADERS.suppliers);
    this._appendSheet(workbook, "purchase_orders", this.purchaseOrders, HEADERS.purchase_orders);
    this._appendSheet(workbook, "purchase_order_items", this.purchaseOrderItems, HEADERS.purchase_order_items);
    this._appendSheet(workbook, "expenses", this.expenses, HEADERS.expenses);
    this._appendSheet(workbook, "daily_closings", this.dailyClosings, HEADERS.daily_closings);
    this._appendSheet(workbook, "supplier_payments", this.supplierPayments, HEADERS.supplier_payments);
    this._appendSheet(workbook, "discounts", this.discounts, HEADERS.discounts);
    this._appendSheet(workbook, "discount_usage", this.discountUsage, HEADERS.discount_usage);
    this._appendSheet(workbook, "inventory_records", this.inventoryRecords, HEADERS.inventory_records);
    this._appendSheet(workbook, "customers", this.members, HEADERS.customers);
    this._appendSheet(workbook, "sales", this._salesRowsForWorkbook(), HEADERS.sales);
    this._appendSheet(workbook, "outstanding_invoices", this.outstandingInvoices, HEADERS.outstanding_invoices);
    this._appendSheet(workbook, "customer_ledger", this.customerLedger, HEADERS.customer_ledger);
    this._appendSheet(workbook, "company_settings", this.companySettings, HEADERS.company_settings);
    this._appendSheet(workbook, "settings", this.settingsRows, HEADERS.settings);
    this._appendSheet(
      workbook,
      "meta",
      Object.entries(this.meta).map(([key, value]) => ({ key, value })),
      HEADERS.meta
    );
    XLSX.writeFile(workbook, this.filePath);
  }

  _appendSheet(workbook, sheetName, data, header) {
    const sheet = XLSX.utils.json_to_sheet(data, { header });
    XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
  }
}

module.exports = {
  ExcelStore
};
