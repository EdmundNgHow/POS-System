CREATE DATABASE IF NOT EXISTS `pos_system`
  CHARACTER SET utf8mb4
  COLLATE utf8mb4_unicode_ci;

USE `pos_system`;

CREATE TABLE IF NOT EXISTS `pos_documents` (
  `collection` VARCHAR(80) NOT NULL,
  `doc_id` VARCHAR(160) NOT NULL,
  `data_json` LONGTEXT NOT NULL,
  `updated_at` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`collection`, `doc_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `pos_sessions` (
  `token_hash` CHAR(64) NOT NULL PRIMARY KEY,
  `user_id` INT NOT NULL,
  `username` VARCHAR(120) NOT NULL,
  `created_at` DATETIME NOT NULL,
  `last_seen_at` DATETIME NOT NULL,
  `expires_at` DATETIME NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
