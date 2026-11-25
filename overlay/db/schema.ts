import { index, integer, primaryKey, sqliteTable, text } from "drizzle-orm/sqlite-core";

export const oneDriveAuth = sqliteTable("onedrive_auth", {
  userId: text("user_id").notNull(),
  kind: text("kind").notNull(),
  revision: text("revision").notNull(),
  clientId: text("client_id").notNull(),
  sealed: text("sealed").notNull(),
  expiresAt: integer("expires_at").notNull(),
  intervalSeconds: integer("interval_s").notNull().default(0),
  nextAt: integer("next_at").notNull().default(0),
}, table => [primaryKey({ columns: [table.userId, table.kind] })]);

export const driveTransfers = sqliteTable("drive_transfers", {
  id: text("id").primaryKey(),
  owner: text("owner").notNull(),
  sessionJson: text("session_json").notNull(),
  expiresAt: integer("expires_at").notNull(),
}, table => [index("drive_transfers_expiry").on(table.expiresAt)]);
