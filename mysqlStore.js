const fs = require("fs");
const path = require("path");
const MySql = require("sync-mysql");
const { ExcelStore } = require("./excelStore");

const COLLECTIONS = [
  ["products", "products"],
  ["orders", "orders"],
  ["order_items", "orderItems"],
  ["payments", "payments"],
  ["pending_orders", "pendingOrders"],
  ["edit_logs", "editLogs"],
  ["users", "users"],
  ["login_history", "loginHistory"],
  ["members", "members"],
  ["suppliers", "suppliers"],
  ["purchase_orders", "purchaseOrders"],
  ["purchase_order_items", "purchaseOrderItems"],
  ["expenses", "expenses"],
  ["daily_closings", "dailyClosings"],
  ["supplier_payments", "supplierPayments"],
  ["discounts", "discounts"],
  ["discount_usage", "discountUsage"],
  ["inventory_records", "inventoryRecords"],
  ["outstanding_invoices", "outstandingInvoices"],
  ["customer_ledger", "customerLedger"],
  ["company_settings", "companySettings"],
  ["settings", "settingsRows"]
];

function sanitizeIdentifier(value, fallback) {
  const text = String(value || "").trim();
  return /^[A-Za-z0-9_$]+$/.test(text) ? text : fallback;
}

function quoteIdentifier(value) {
  return `\`${sanitizeIdentifier(value, "pos_system").replace(/`/g, "``")}\``;
}

function timestampForFile(value = new Date()) {
  const date = value instanceof Date ? value : new Date(value);
  const pad = (part) => String(part).padStart(2, "0");
  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
    "-",
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds())
  ].join("");
}

class MySqlStore extends ExcelStore {
  constructor({
    host = "127.0.0.1",
    port = 3306,
    user = "root",
    password = "",
    database = "pos_system",
    tablePrefix = "pos_",
    legacyExcelFile = "",
    backupDir = ""
  } = {}) {
    super(legacyExcelFile || path.join(process.cwd(), "data", "pos-data.xlsx"));
    this.storageType = "mysql";
    this.config = {
      host,
      port: Number(port || 3306),
      user,
      password,
      database: sanitizeIdentifier(database, "pos_system"),
      tablePrefix: sanitizeIdentifier(tablePrefix, "pos_")
    };
    this.legacyExcelFile = legacyExcelFile || "";
    this.backupDir = backupDir || path.join(process.cwd(), "backups");
    this.connection = null;
    this.lastReconnectAt = "";
    this.documentsTable = `${this.config.tablePrefix}documents`;
    this.sessionsTable = `${this.config.tablePrefix}sessions`;
  }

  initialize() {
    this._connect();
    this._ensureSchema();
    if (!this._hasDocuments()) {
      this._seedDatabase();
    }
    this.reloadFromDatabase();
  }

  _connect() {
    const admin = new MySql({
      host: this.config.host,
      port: this.config.port,
      user: this.config.user,
      password: this.config.password,
      multipleStatements: true
    });
    admin.query(`CREATE DATABASE IF NOT EXISTS ${quoteIdentifier(this.config.database)} CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci`);
    this.connection = new MySql({
      host: this.config.host,
      port: this.config.port,
      user: this.config.user,
      password: this.config.password,
      database: this.config.database,
      multipleStatements: true
    });
    this.lastReconnectAt = new Date().toISOString();
  }

  _isConnectionError(err) {
    const code = String(err?.code || "");
    const message = String(err?.message || err || "");
    return [
      "PROTOCOL_CONNECTION_LOST",
      "ECONNRESET",
      "ECONNREFUSED",
      "ETIMEDOUT",
      "EPIPE"
    ].includes(code) || /connection.*(lost|closed|refused|reset)|cannot enqueue|socket|timeout/i.test(message);
  }

  _query(sql, params = [], options = {}) {
    const allowReconnect = options.allowReconnect !== false;
    try {
      if (!this.connection) {
        this._connect();
      }
      return this.connection.query(sql, params);
    } catch (err) {
      if (!allowReconnect || !this._isConnectionError(err)) {
        throw err;
      }
      this.connection = null;
      this._connect();
      return this.connection.query(sql, params);
    }
  }

  _runTransaction(work, allowRetry = true) {
    this._query("START TRANSACTION");
    try {
      const result = work();
      this._query("COMMIT", [], { allowReconnect: false });
      return result;
    } catch (err) {
      try {
        this._query("ROLLBACK", [], { allowReconnect: false });
      } catch (_) {
        // Preserve original transaction failure.
      }
      if (allowRetry && this._isConnectionError(err)) {
        this.connection = null;
        this._connect();
        return this._runTransaction(work, false);
      }
      throw err;
    }
  }

  _ensureSchema() {
    this._query(
      `CREATE TABLE IF NOT EXISTS ${quoteIdentifier(this.documentsTable)} (
        collection VARCHAR(80) NOT NULL,
        doc_id VARCHAR(160) NOT NULL,
        data_json LONGTEXT NOT NULL,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        PRIMARY KEY (collection, doc_id)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci`
    );
    this._query(
      `CREATE TABLE IF NOT EXISTS ${quoteIdentifier(this.sessionsTable)} (
        token_hash CHAR(64) NOT NULL PRIMARY KEY,
        user_id INT NOT NULL,
        username VARCHAR(120) NOT NULL,
        created_at DATETIME NOT NULL,
        last_seen_at DATETIME NOT NULL,
        expires_at DATETIME NOT NULL
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci`
    );
  }

  _hasDocuments() {
    const rows = this._query(`SELECT COUNT(*) AS count FROM ${quoteIdentifier(this.documentsTable)}`);
    return Number(rows?.[0]?.count || 0) > 0;
  }

  _seedDatabase() {
    if (this.legacyExcelFile && fs.existsSync(this.legacyExcelFile)) {
      const legacy = new ExcelStore(this.legacyExcelFile);
      legacy.initialize();
      for (const [, propertyName] of COLLECTIONS) {
        this[propertyName] = Array.isArray(legacy[propertyName]) ? legacy[propertyName].map((row) => ({ ...row })) : [];
      }
      this.meta = { ...(legacy.meta || {}) };
    } else {
      this._seedDefaults();
      this._ensureCompanySettingsSeeded();
      this._ensureSettingsSeeded();
    }
    this._persistToDatabase();
  }

  _resetRuntimeData() {
    for (const [, propertyName] of COLLECTIONS) {
      this[propertyName] = [];
    }
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

  reloadFromDatabase() {
    this._resetRuntimeData();
    const rows = this._query(
      `SELECT collection, doc_id, data_json FROM ${quoteIdentifier(this.documentsTable)}
       ORDER BY collection, CAST(doc_id AS UNSIGNED), doc_id`
    );
    const propertyByCollection = new Map(COLLECTIONS);
    for (const row of rows) {
      const collection = String(row.collection || "");
      const parsed = JSON.parse(String(row.data_json || "{}"));
      if (collection === "meta") {
        this.meta[String(row.doc_id || "")] = Number(parsed.value || 0);
        continue;
      }
      const propertyName = propertyByCollection.get(collection);
      if (!propertyName) {
        continue;
      }
      this[propertyName].push(parsed);
    }
    this._ensureUsersSeeded();
    this._ensureDiscountsSeeded();
    const companyChanged = this._ensureCompanySettingsSeeded();
    this._ensureSettingsSeeded();
    this._rebuildMeta();
    if (companyChanged) {
      this._persistToDatabase();
    }
  }

  _rowDocId(row, fallbackIndex) {
    if (row && Object.hasOwn(row, "id") && row.id !== undefined && row.id !== null && row.id !== "") {
      return String(row.id);
    }
    if (row && Object.hasOwn(row, "key") && row.key) {
      return String(row.key);
    }
    return String(fallbackIndex + 1);
  }

  _persistToDatabase() {
    this._runTransaction(() => {
      this._query(`DELETE FROM ${quoteIdentifier(this.documentsTable)}`, [], { allowReconnect: false });
      for (const [collection, propertyName] of COLLECTIONS) {
        const rows = Array.isArray(this[propertyName]) ? this[propertyName] : [];
        rows.forEach((row, index) => {
          this._query(
            `INSERT INTO ${quoteIdentifier(this.documentsTable)} (collection, doc_id, data_json) VALUES (?, ?, ?)`,
            [collection, this._rowDocId(row, index), JSON.stringify(row)],
            { allowReconnect: false }
          );
        });
      }
      for (const [key, value] of Object.entries(this.meta || {})) {
        this._query(
          `INSERT INTO ${quoteIdentifier(this.documentsTable)} (collection, doc_id, data_json) VALUES (?, ?, ?)`,
          ["meta", key, JSON.stringify({ value })],
          { allowReconnect: false }
        );
      }
    });
  }

  _save() {
    this._ensureCompanySettingsSeeded();
    this._ensureSettingsSeeded();
    this._rebuildMeta();
    this._persistToDatabase();
  }

  checkConnection() {
    this._query("SELECT 1 AS ok");
    return true;
  }

  getStorageInfo() {
    return {
      type: "mysql",
      host: this.config.host,
      port: this.config.port,
      database: this.config.database,
      documents_table: this.documentsTable,
      sessions_table: this.sessionsTable,
      backup_dir: this.backupDir,
      last_reconnect_at: this.lastReconnectAt
    };
  }

  saveSession(tokenHash, session) {
    this._query(
      `REPLACE INTO ${quoteIdentifier(this.sessionsTable)}
       (token_hash, user_id, username, created_at, last_seen_at, expires_at)
       VALUES (?, ?, ?, ?, ?, ?)`,
      [
        tokenHash,
        Number(session.user_id || 0),
        String(session.username || ""),
        String(session.created_at || new Date().toISOString()).slice(0, 19).replace("T", " "),
        String(session.last_seen_at || new Date().toISOString()).slice(0, 19).replace("T", " "),
        String(session.expires_at || new Date(Date.now() + 12 * 60 * 60 * 1000).toISOString()).slice(0, 19).replace("T", " ")
      ]
    );
  }

  getSession(tokenHash) {
    const rows = this._query(`SELECT * FROM ${quoteIdentifier(this.sessionsTable)} WHERE token_hash = ? LIMIT 1`, [tokenHash]);
    const row = rows?.[0];
    if (!row) return null;
    return {
      user_id: Number(row.user_id || 0),
      username: String(row.username || ""),
      created_at: new Date(row.created_at).toISOString(),
      last_seen_at: new Date(row.last_seen_at).toISOString(),
      expires_at: new Date(row.expires_at).toISOString()
    };
  }

  deleteSession(tokenHash) {
    this._query(`DELETE FROM ${quoteIdentifier(this.sessionsTable)} WHERE token_hash = ?`, [tokenHash]);
  }

  cleanupExpiredSessions(now = new Date()) {
    const value = now.toISOString().slice(0, 19).replace("T", " ");
    this._query(`DELETE FROM ${quoteIdentifier(this.sessionsTable)} WHERE expires_at < ?`, [value]);
  }

  createBackup(actor = "system") {
    fs.mkdirSync(this.backupDir, { recursive: true });
    const rows = this._query(`SELECT collection, doc_id, data_json, updated_at FROM ${quoteIdentifier(this.documentsTable)} ORDER BY collection, doc_id`);
    const backup = {
      generated_at: new Date().toISOString(),
      generated_by: String(actor || "system"),
      storage: this.getStorageInfo(),
      rows
    };
    const fileName = `pos-backup-${timestampForFile()}.json`;
    const filePath = path.join(this.backupDir, fileName);
    fs.writeFileSync(filePath, JSON.stringify(backup, null, 2), "utf8");
    return { file_name: fileName, file_path: filePath, generated_at: backup.generated_at, row_count: rows.length };
  }

  listBackups() {
    fs.mkdirSync(this.backupDir, { recursive: true });
    return fs.readdirSync(this.backupDir)
      .filter((fileName) => /^pos-backup-\d{8}-\d{6}\.json$/.test(fileName))
      .map((fileName) => {
        const filePath = path.join(this.backupDir, fileName);
        const stats = fs.statSync(filePath);
        return {
          file_name: fileName,
          file_path: filePath,
          size_bytes: stats.size,
          created_at: stats.birthtime.toISOString(),
          updated_at: stats.mtime.toISOString()
        };
      })
      .sort((a, b) => String(b.file_name).localeCompare(String(a.file_name)));
  }

  restoreBackup(fileName, actor = "system") {
    const safeName = path.basename(String(fileName || ""));
    const filePath = path.join(this.backupDir, safeName);
    const resolved = path.resolve(filePath);
    const backupRoot = path.resolve(this.backupDir);
    if (!(resolved === backupRoot || resolved.startsWith(`${backupRoot}${path.sep}`)) || !fs.existsSync(resolved)) {
      const err = new Error("Backup file not found.");
      err.code = "NOT_FOUND";
      throw err;
    }

    const backup = JSON.parse(fs.readFileSync(resolved, "utf8"));
    if (!Array.isArray(backup.rows)) {
      const err = new Error("Invalid backup file.");
      err.code = "BAD_REQUEST";
      throw err;
    }

    this._runTransaction(() => {
      this._query(`DELETE FROM ${quoteIdentifier(this.documentsTable)}`, [], { allowReconnect: false });
      for (const row of backup.rows) {
        this._query(
          `INSERT INTO ${quoteIdentifier(this.documentsTable)} (collection, doc_id, data_json) VALUES (?, ?, ?)`,
          [row.collection, row.doc_id, row.data_json],
          { allowReconnect: false }
        );
      }
    });
    this.reloadFromDatabase();
    this.logAudit({
      action: "database_restore",
      module: "System",
      changedBy: actor,
      description: `Restored database backup ${safeName}`,
      details: { file_name: safeName }
    });
    return { file_name: safeName, restored_at: new Date().toISOString() };
  }
}

module.exports = {
  MySqlStore
};
