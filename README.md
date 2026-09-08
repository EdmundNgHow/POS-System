# Tunas Badminton Sport Centre POS

Badminton centre management POS system with:
- Frontend: HTML + CSS + vanilla JavaScript
- Backend: Node.js + Express
- Database: MySQL/MariaDB for production (`POS_STORAGE=mysql`); local auto mode falls back to legacy storage until DB config is added
- Desktop build: Electron (Windows `.exe`)

## Architecture
- RMS module definitions live in `server/modules/definitions.js`
- Role/permission grants live in `server/security/permissions.js`
- Bootstrap services live in `server/bootstrap/`
- Frontend navigation is driven by `GET /api/modules`
- Full architecture guide: `docs/architecture.md`

## Web Run (Local)
1. `npm install`
2. For production, create a MySQL/MariaDB database/user, or run `db/schema.mysql.sql`
3. For production, set `POS_STORAGE=mysql`, `MYSQL_HOST`, `MYSQL_USER`, `MYSQL_PASSWORD`, and `MYSQL_DATABASE`
4. Run `npm run db:init` to create/verify tables and seed existing local data into MySQL/MariaDB
5. For local testing without MySQL, leave `POS_STORAGE=auto`
6. `npm start`
7. Open `http://localhost:3000/login`

If `POS_STORAGE=mysql` is set, startup intentionally fails when MySQL/MariaDB is unavailable. Excel/local legacy mode is not production storage.

## Default System Accounts
- System Owner: username `TSSC`, temporary password `Tssc@123`, forced to change password on first login
- Super Admin: username `superadmin`, temporary password `Super@123`, forced to change password on first login
- Admin: username `admin`, temporary password `Admin@123`, forced to change password on first login

## Desktop Run (Local)
1. `npm install`
2. `npm run desktop`

## LAN Access (Phones on Same WiFi)
1. Start the desktop app with `npm run desktop` or the built `.exe`.
2. Open the dashboard and find the `Mobile POS Access` panel.
3. Connect phones to the same WiFi network, then scan the QR code or open the displayed mobile URL.
4. Sign in with an admin or staff account. Phones use the same Express server and database as the desktop POS, so sales, stock, members, and reports stay synchronized.
5. Production devices should all connect to the same Express backend URL, backed by the same MySQL/MariaDB database.

## Build Windows EXE (Portable)
1. `npm install`
2. `npm run build:exe`
3. EXE output folder: `dist/Tunas Badminton Sport Centre-win32-x64/`

## Build Windows Installer (NSIS)
1. `npm install`
2. `npm run build:win`
3. Installer output will be in `dist/`

## Main Pages
- Login: `/login`
- Dashboard: `/dashboard`
- Sales POS: `/sales`
- Receipt History: `/receipts`
- Outstanding Orders: `/outstanding`
- Customer Ledger: `/customer-ledger` or `/reports/customer-ledger`
- Today's Sold Items: `/todays-sold-items` or `/reports/todays-sold-items`
- Staff Management: `/staff` (admin)
- Discount Management: `/discounts` (admin)
- Item Management: `/manage` (admin and staff)
- Member Management: `/members`
- Reports: `/reports` (admin)
- Audit Log: `/audit-log` (admin)
- Settings: `/settings` (admin)
- User Management: `/settings/users` (admin / super admin)
- Legacy Stock Management: `/stock`
- Backward-compatible aliases: `/landing`, `/history`, `/summary`, `/edit-log`, `/staff`

## Features
- Role-based login (admin / staff)
- Dashboard with inventory alerts, outstanding orders, today's unpaid sales, outstanding amount, overdue orders, and recent payments
- Dashboard LAN IP display and QR code for same-WiFi mobile access
- Express backend listens on `0.0.0.0`, so phones/tablets can use `http://<LOCAL_IP>:3000`
- Realtime server-sent events notify connected devices when data changes; pages refresh their data without manual reloads
- Sales POS with always-ready barcode scanner focus, member lookup, approved discounts, cart totals, multiple persisted active orders, descriptions, and receipt printing
- Sales POS can add a selected customer's total outstanding balance to the cart as one `Outstanding Payment` line and collect it through checkout
- Multiple order tabs let cashiers create, switch, save, resume, cancel, and complete active orders without losing carts after restart
- Payment methods include Cash, DuitNow QR, Credit Card, Debit Card, Bank Transfer, E-Wallet, and Haven't Paid (Belum Bayar)
- Admin Discount Management with unlimited percentage, fixed, member, product, category, and promotional discount rules
- Staff Apply Discount flow that lists active approved discounts and sends only the saved rule ID to checkout
- Discount usage tracking with staff name, transaction number, date/time, and discount amount
- Dedicated item management page for admin/staff item CRUD, with no popup-only inventory management
- Item search, category/brand/supplier/stock/status filters, sorting, pagination, stock adjustment, details panel, barcode labels, and badminton item templates
- Admin can view cost, profit, margin, reports, staff management, audit logs, settings, and discount rule management APIs
- Staff can manage items and members, process sales, and apply approved/member discounts
- Staff cannot view product cost, margin, admin reports, audit logs, settings, receipt void/refund/delete, or user/discount management APIs
- User Management supports Super Admin, Admin, Manager, Cashier, and Inventory Staff roles, user status control, password reset, login history, and configurable role permissions
- New passwords are stored as salted PBKDF2 hashes; legacy SHA-256 password hashes remain readable for existing accounts
- Member management with tier discounts
- Outstanding Orders page lists unpaid/partially paid receipts, supports search/filter, collects later payments, and prints updated receipts
- Customer Ledger records outstanding created as Credit, customer payments as Debit, and calculates running balance as Credit minus Debit
- MySQL/MariaDB is the production source of truth for customers, products, sales, outstanding invoices, customer ledger, payments, settings, sessions, and audit logs
- Excel is retained for exports/reports/backups only and must not be used as the production database
- Database backups can be created manually or automatically as JSON backup snapshots in `backups/`
- MySQL/MariaDB mode automatically reconnects after dropped database connections and keeps page users warned through the global health banner
- Today's Sold Items report lists sold items, category/product/top-selling summaries, low-stock warnings, payment summary, and staff sales summary with automatic live refresh after sales
- Admins can export Today's Sold Items to PDF/Excel and print reports; staff can view today's sold items without cost, profit, or margin
- Receipt history includes payment status, outstanding balance, description, payment method, search, and status filtering
- Receipts display payment status and outstanding balance
- Discount rules for approved percentage/fixed, quantity, member, product, category, and promotion discounts
- Inventory records for item creation, stock adjustment, and sales movement
- Admin receipt editing with bill line editing and delete; staff payment-status updates for authorised statuses
- Receipt rounding (cash)
- Daily summary with profit for admin
- Dedicated stock page with search and add-stock action
- Audit log with login/logout, item, order created/completed/cancelled/marked unpaid, outstanding payment collection, payment status, member, staff, and discount change tracking
- Electron desktop mode disables renderer background throttling and blocks app suspension while the POS is running
- MYR currency display

## Security Note
- Credentials are not displayed in the login UI.
- Users should sign in using assigned accounts.
- Seed accounts for a fresh installation are `superadmin` / `Super@123` and `admin` / `Admin@123`.
- Default Super Admin and Admin accounts must change their password on first login before accessing the system.
- Staff accounts are not created automatically; create them from Settings → User Management.

## API (Core)
- `POST /api/auth/login`
- `POST /api/auth/logout`
- `GET /api/auth/me`
- `GET /api/modules`
- `GET /api/users` (admin)
- `POST /api/users` (admin)
- `GET /api/users/:id` (admin)
- `PUT /api/users/:id` (admin)
- `DELETE /api/users/:id` (super admin)
- `POST /api/users/:id/reset-password` (admin)
- `GET /api/users/login-history` (audit permission)
- `GET /api/roles` (admin)
- `PUT /api/roles/:role/permissions` (admin)
- `GET /api/products`
- `GET /api/product-filters`
- `POST /api/products`
- `PUT /api/products/:id`
- `DELETE /api/products/:id`
- `GET /api/members`
- `POST /api/members`
- `PUT /api/members/:id`
- `DELETE /api/members/:id`
- `GET /api/discounts`
- `POST /api/discounts` (admin)
- `PUT /api/discounts/:id` (admin)
- `DELETE /api/discounts/:id` (admin)
- `GET /api/discount-usage` (admin)
- `GET /api/stock`
- `POST /api/stock/add`
- `GET /api/inventory/alerts`
- `GET /api/inventory/records`
- `GET /api/pending-orders`
- `POST /api/pending-orders`
- `DELETE /api/pending-orders/:id`
- `POST /api/orders`
- `GET /api/orders/:id`
- `POST /api/orders/:id/collect-payment`
- `PUT /api/orders/:id`
- `DELETE /api/orders/:id` (admin)
- `GET /api/orders`
- `GET /api/outstanding-orders`
- `GET /api/outstanding-summary`
- `GET /api/customers/outstanding`
- `GET /api/customer-ledger`
- `GET /api/reports/customer-ledger`
- `GET /api/edit-logs` (admin)
- `GET /api/admin/summary` (admin)
- `GET /api/sales/summary` (admin)
- `GET /api/reports/todays-sold-items`
- `GET /api/reports/todays-sold-items/export.xlsx` (admin)
- `GET /api/reports/todays-sold-items/export.pdf` (admin)
- `GET /api/events`
- `GET /api/backups` (admin)
- `POST /api/backups` (admin)
- `POST /api/backups/restore` (admin)
- `GET /api/settings` (admin)
- `GET /api/network/access`
- `GET /api/health`
