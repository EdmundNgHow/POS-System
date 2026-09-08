# POS Completion Checklist

This checklist records the current project audit before adding more business features. The goal is to complete the system without removing existing routes, APIs, data, reports, or working screens.

## Completed Or Mostly Implemented

- [x] **Authentication and sessions**: username/password login, persistent server sessions, logout, forced password change for default accounts.
- [x] **Default authorities**: System Owner, Super Admin, and Admin roles exist with protected System Owner handling.
- [x] **RBAC foundation**: module registry, route protection, permission defaults, role permission editing, and dynamic sidebar visibility.
- [x] **Primary navigation cleanup**: primary sidebar now focuses on Dashboard, Sales, Inventory, Members, Reports, and Settings.
- [x] **Sales POS basics**: product search, barcode entry, cart, discounts, checkout, pending orders, multiple active order UI, receipts, and outstanding payment collection.
- [x] **Outstanding receivables**: unpaid/partial transactions, outstanding summary, customer ledger, payment collection, and receipt status updates.
- [x] **Inventory basics**: item CRUD, barcode/SKU/category/brand/supplier fields, stock adjustment, inventory records, low-stock alerts, and barcode label printing.
- [x] **Members basics**: member CRUD, member lookup in sales, customer linkage for orders, and outstanding balance visibility.
- [x] **Discounts and promotions**: discount CRUD, active discount selection, minimum purchase rules, and discount usage reporting.
- [x] **Receipt center basics**: receipt history, reprint, refund/void/status actions, payment method and payment status display.
- [x] **Reports restoration**: Reports route restored, Today's Sold Items available, modern report dashboard shell, filters, charts, exports, and live refresh hooks.
- [x] **Settings consolidation**: one Settings Center page with users, roles, discounts, backups, database status, printer/template placeholders, and system overview.
- [x] **Audit logging basics**: user login/logout and business actions are logged to the audit/edit log.
- [x] **Network access**: Express server listens on all interfaces and exposes local network URL for mobile/tablet access.
- [x] **Storage adapters**: Excel fallback storage exists; MySQL document-store adapter and initialization script exist.
- [x] **Backup/restore basics**: manual backup and restore endpoints are available for the active storage adapter.
- [x] **Live update channel**: Server-Sent Events endpoint exists and UI pages listen for business updates.

## Incomplete Or Partial Features

- [x] **Purchasing module**: real Purchasing page, purchase order API, stock receiving, purchase history, stock movement logging, and supplier payable creation are implemented.
- [x] **Suppliers module**: real Suppliers page, supplier CRUD API, contacts, payment terms, and outstanding payable tracking are implemented.
- [x] **Finance module foundation**: `/finance` now opens a Finance Center with summary, expenses, outstanding payables, supplier payments, and daily closing.
- [ ] **Finance module advanced controls**: cash drawer sessions, payable aging, and richer income/expense reporting still need deeper workflow polish.
- [ ] **Members advanced features**: purchase history UI, reward points, vouchers, membership tiers, and redemption rules are not complete.
- [ ] **Settings functional depth**: printer settings, receipt templates, database settings, and system preferences mostly show guidance/placeholders instead of full editable forms.
- [ ] **Custom roles**: role permissions can be edited, but creating/deleting future custom roles is not fully implemented.
- [ ] **Report coverage**: redesigned dashboard exists, but many category links still reuse existing pages or anchors rather than dedicated report screens.
- [ ] **Report exports**: Today's Sold Items supports Excel/PDF export; full PDF/Excel/CSV export coverage for every report category is incomplete.
- [ ] **Profit/cost reporting**: permission hooks exist, but profit, margin, and cost views need stronger report-level implementation.
- [ ] **Offline transaction queue**: UI warns/reconnects, but unsaved offline sales are not queued and replayed after database recovery.
- [ ] **Normalized MySQL migrations**: MySQL uses `pos_documents`; production-grade normalized tables for every business entity are not yet implemented.
- [ ] **Multi-device conflict handling**: real-time refresh exists, but optimistic conflict checks and server-side versioning are limited.
- [ ] **Printer integration**: receipt printing uses browser/Electron print flow; dedicated printer profiles and template controls are incomplete.
- [ ] **Mobile/iPad polish**: responsive CSS exists, but high-density sales, reports, and management tables need focused tablet QA.

## Duplicated Or Legacy UI

- [x] **Duplicate Settings sidebar entries**: resolved at the primary navigation level.
- [ ] **Legacy header links**: older pages still contain page-local navigation links that are normalized by `nav.js`; they should be gradually simplified to reduce clutter.
- [x] **Placeholder modules**: Purchasing and Suppliers no longer use `module-placeholder.html`.
- [x] **Finance overlap**: `/finance` now has its own Finance Center while receivables remain available through `/finance/outstanding-receivables`.
- [ ] **Report overlap**: several report links route to existing stock/history/outstanding pages rather than purpose-built report tables.

## Missing Pages

- [x] `suppliers.html` and `suppliers.js`
- [x] `purchasing.html` and `purchasing.js`
- [x] `finance.html` and `finance.js`
- [x] Daily Closing section in Finance Center
- [x] Expenses section in Finance Center
- [x] Outstanding Payables section in Finance Center
- [x] Supplier Payments section in Finance Center
- [ ] Reward Points / Vouchers page or members tab
- [ ] Receipt Template editor
- [ ] Printer Management editor
- [ ] Database Settings editor
- [ ] Custom Role editor

## Missing Or Needed API Endpoints

- [x] `GET /api/suppliers`
- [x] `POST /api/suppliers`
- [x] `PUT /api/suppliers/:id`
- [x] `DELETE /api/suppliers/:id`
- [x] `GET /api/purchasing/orders`
- [x] `POST /api/purchasing/orders`
- [x] `GET /api/purchasing/orders/:id`
- [x] `POST /api/purchasing/orders/:id/receive`
- [x] `GET /api/purchasing/history`
- [x] `GET /api/finance/summary`
- [x] `GET /api/finance/expenses`
- [x] `POST /api/finance/expenses`
- [x] `GET /api/finance/daily-closing`
- [x] `POST /api/finance/daily-closing`
- [x] `GET /api/finance/payables`
- [x] `POST /api/finance/payables/:supplierId/payment`
- [ ] `GET /api/members/:id/history`
- [ ] `GET /api/members/:id/rewards`
- [ ] `POST /api/members/:id/rewards/adjust`
- [ ] `GET /api/vouchers`
- [ ] `POST /api/vouchers`
- [ ] `GET /api/settings/printer`
- [ ] `PUT /api/settings/printer`
- [ ] `GET /api/settings/receipt-template`
- [ ] `PUT /api/settings/receipt-template`

## Missing Database Collections Or Tables

- [x] `suppliers`
- [x] `purchase_orders`
- [x] `purchase_order_items`
- [x] `supplier_payments`
- [x] `expenses`
- [x] `daily_closings`
- [ ] `cash_drawer_sessions`
- [ ] `member_reward_ledger`
- [ ] `vouchers`
- [ ] `voucher_redemptions`
- [ ] `printer_settings`
- [ ] `receipt_templates`
- [ ] `offline_queue`
- [ ] Normalized MySQL tables with indexes for products, orders, order items, payments, stock records, users, members, suppliers, purchases, finance, and audit logs.

## UI Inconsistencies To Clean Up

- [ ] Older pages use different header link sets and labels.
- [ ] Some management tables are dense on tablet width.
- [ ] Settings cards mix implemented actions with placeholder guidance.
- [ ] Report category cards link to mixed destinations: pages, anchors, and reused modules.
- [ ] Admin-only links on older pages still rely on legacy role checks in addition to RBAC navigation.
- [ ] Some tables render `0` values as blank/fallback text in older render helpers.

## Performance And Stability Risks

- [ ] Excel fallback rewrites the workbook on many mutations; this is acceptable for small stores but not production scale.
- [ ] MySQL adapter stores JSON documents and reloads collections instead of using indexed relational queries.
- [ ] Report aggregation often scans full in-memory arrays.
- [ ] Product search should use indexed database queries for large inventories.
- [ ] Barcode scanning relies on browser focus recovery; needs long-idle and reconnect QA on the target scanners.
- [ ] Server uses synchronous MySQL access, which can block under heavy multi-device load.
- [ ] Some pages reload full lists after every mutation instead of applying targeted updates.
- [ ] Offline recovery does not yet preserve unsaved transactions if the database drops mid-sale.

## Route Status

- [x] `/dashboard`
- [x] `/sales`
- [x] `/inventory`
- [x] `/manage`
- [x] `/stock`
- [x] `/members`
- [x] `/reports`
- [x] `/todays-sold-items`
- [x] `/receipts`
- [x] `/history`
- [x] `/outstanding`
- [x] `/customer-ledger`
- [x] `/audit`
- [x] `/settings`
- [x] `/settings/users`
- [x] `/purchasing`
- [x] `/suppliers`
- [x] `/finance`

## Recommended Implementation Order

1. [x] Replace Supplier placeholder with persistent supplier CRUD.
2. [x] Replace Purchasing placeholder with purchase orders, receiving, and purchase history.
3. [x] Add Finance Center for receivables, payables, expenses, daily closing, and income summary.
4. [ ] Make printer settings and receipt templates editable in Settings.
5. [ ] Add member reward points, vouchers, and member purchase history.
6. [ ] Expand report exports and dedicated report tables.
7. [ ] Add offline transaction queue and replay safeguards.
8. [ ] Move from JSON document MySQL storage toward indexed normalized tables.
9. [ ] Perform tablet/browser/Electron scanner endurance testing.
10. [ ] Run full regression validation across sales, inventory, receipts, reports, RBAC, and backups.
