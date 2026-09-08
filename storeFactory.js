const { ExcelStore } = require("./excelStore");
const { MySqlStore } = require("./mysqlStore");

function env(name, fallback = "") {
  return process.env[name] === undefined ? fallback : process.env[name];
}

function hasDatabaseConfig() {
  return [
    "MYSQL_HOST",
    "MYSQL_PORT",
    "MYSQL_USER",
    "MYSQL_PASSWORD",
    "MYSQL_DATABASE",
    "DB_HOST",
    "DB_PORT",
    "DB_USER",
    "DB_PASSWORD",
    "DB_NAME"
  ].some((key) => String(process.env[key] || "").trim());
}

function createStore({ legacyExcelFile, backupDir } = {}) {
  const requestedStorage = String(env("POS_STORAGE", env("DB_CLIENT", "auto"))).trim().toLowerCase();
  const storage = requestedStorage === "auto"
    ? hasDatabaseConfig()
      ? "mysql"
      : "excel"
    : requestedStorage;
  if (storage === "excel" || storage === "xlsx") {
    const store = new ExcelStore(legacyExcelFile);
    store.storageType = "excel";
    store.getStorageInfo = () => ({
      type: "excel",
      file_path: legacyExcelFile,
      warning: "Legacy local mode. Configure MySQL/MariaDB for production multi-device storage."
    });
    return store;
  }

  return new MySqlStore({
    host: env("MYSQL_HOST", env("DB_HOST", "127.0.0.1")),
    port: Number(env("MYSQL_PORT", env("DB_PORT", "3306"))),
    user: env("MYSQL_USER", env("DB_USER", "root")),
    password: env("MYSQL_PASSWORD", env("DB_PASSWORD", "")),
    database: env("MYSQL_DATABASE", env("DB_NAME", "pos_system")),
    tablePrefix: env("MYSQL_TABLE_PREFIX", "pos_"),
    legacyExcelFile,
    backupDir
  });
}

module.exports = {
  createStore
};
