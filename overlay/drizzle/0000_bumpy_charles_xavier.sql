CREATE TABLE `drive_transfers` (
	`id` text PRIMARY KEY NOT NULL,
	`owner` text NOT NULL,
	`session_json` text NOT NULL,
	`expires_at` integer NOT NULL
);
--> statement-breakpoint
CREATE INDEX `drive_transfers_expiry` ON `drive_transfers` (`expires_at`);--> statement-breakpoint
CREATE TABLE `onedrive_auth` (
	`user_id` text NOT NULL,
	`kind` text NOT NULL,
	`revision` text NOT NULL,
	`client_id` text NOT NULL,
	`sealed` text NOT NULL,
	`expires_at` integer NOT NULL,
	`interval_s` integer DEFAULT 0 NOT NULL,
	`next_at` integer DEFAULT 0 NOT NULL,
	PRIMARY KEY(`user_id`, `kind`)
);
