# Production RMS Architecture

This project is being refactored from a POS-only app into the Tunas Badminton Sport Centre Retail Management System without rebuilding from scratch.

## Architecture Rules

- Keep all business data in the shared store layer (`db/`) and production MySQL/MariaDB configuration.
- Keep feature ownership in modules under `server/modules`.
- Register page routes from module metadata instead of scattering page routes through `server.js`.
- Register navigation from `/api/modules` so desktop, browser, phones, tablets, and PWA shells share the same module menu.
- Use permission IDs instead of hardcoded page visibility whenever a route represents business functionality.
- Preserve legacy URLs while new module URLs are introduced.

## Backend Layout

- `server/bootstrap/` starts environment, storage, and network services.
- `server/modules/definitions.js` is the source of truth for RMS modules, pages, API base paths, features, and permission IDs.
- `server/modules/registry.js` turns module metadata into page routes and role-aware navigation.
- `server/modules/api.js` exposes the module catalog to clients.
- `server/security/permissions.js` defines roles and permission grants.
- `db/` remains the shared database operation boundary for all modules.

## Modules

- Dashboard: business overview, alerts, performance, notifications.
- Sales: POS, barcode scanning, carts, multiple orders, checkout, receipts.
- Inventory: items, categories, brands, stock movement, low stock.
- Purchasing: suppliers, purchase orders, receive stock, supplier payments.
- Members: registration, purchase history, balances, rewards.
- Suppliers: contacts, supply records, payment history.
- Reports: sales, inventory, product, staff, payment, and financial reporting.
- Finance: receivables, payables, daily closing, cash drawer, expenses.
- Receipt Center: receipt history, reprint, refund, void, payment status.
- Audit: user activity, inventory logs, receipt logs, payment logs.
- Settings: users, roles, permissions, printers, templates, backups, database configuration.

## Migration Rule

Existing handlers in `server.js` remain active until each feature is moved into its module service/router. New work should add module-owned services and routes first, then remove old monolith code only after a smoke test confirms the legacy URL still behaves correctly.
