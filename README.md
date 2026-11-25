# Personal Drive for ChatGPT

This independent, private-use project lets you use ChatGPT with your personal Microsoft OneDrive. Search your drive, work with files and folders, and transfer photos and videos from a copy you install under your own ChatGPT account. It contains the OneDrive application, database schema and tests as an overlay for the ChatGPT Sites Vinext starter.

## What the tools do

| Task | Behavior |
| --- | --- |
| Browse and search | List folders and search drive items with pagination. Read metadata, text, supported Office PDF renditions, image previews and video thumbnails. |
| Create and edit | Create folders and files, replace file contents, rename items and move them within your drive. Updates and deletes require the current item version (eTag). |
| Delete | Move an item to the OneDrive recycle bin. Deleting a folder also recycles its contents. |
| Transfer | Inline files are limited to 1 MiB. The two chat attachment tools accept uploads up to 20 MiB. The private browser page supports larger uploads through Microsoft Graph upload sessions. A private download link requires the same signed-in account. |

## Start with your own private Site

You need access to ChatGPT Sites with MCP plugin publishing, a personal Microsoft account, and access to Microsoft Entra app registrations. The local test suite uses Node.js 22.13 or later. This overlay was tested against the Vinext starter shipped with Sites 1.1.6 and the Miniflare version locked by its Wrangler dependency.

1. In ChatGPT, create a fresh **owner-only** Vinext Site from the standard Sites starter. Have Sites initialize or open its managed source checkout. Keep the starter's files and the new Site's own project identity.
2. Download or clone this source package into a separate directory. In a terminal, run the command below from the package directory, substituting the path of the fresh starter checkout. An alternative in ChatGPT Work Mode is to attach this package and direct Sites to apply it to the new private Site.

   ```sh
   node apply-overlay.mjs "/path/to/new-sites-checkout"
   ```

   The script copies the application into the starter, sets its D1 binding to `DB`, enables `mcp`, and records the starter's locked Miniflare version as a direct test dependency. It checks for existing source changes before writing. It can be rerun on an unchanged installation. If a newer template changes the starter files or its locked runtime, review those differences before porting the overlay.
3. From the **Sites checkout**, install dependencies, run the tests and build.

   ```sh
   npm run install:ci
   node --test
   npm run build
   ```

   The tests use synthetic accounts and files. The Worker regression tests exercise the actual Worker runtime with a simulated Microsoft endpoint. These checks require neither a Microsoft sign-in nor access to a real drive.
4. In Sites runtime settings, set the following values for this new deployment. Use the hosting service's secret mechanism for the key.

   | Setting | Value |
   | --- | --- |
   | `DRIVE_ENCRYPTION_KEY` | A fresh 32-byte cryptographically random key encoded in standard base64. Keep the key stable across deployments. It encrypts per-user Microsoft tokens and upload session credentials. |
   | `DRIVE_SITE_ORIGIN` | Your new Site's canonical HTTPS origin, with no path or trailing slash. The MCP tools use it for connection and transfer links. |

   Sites can generate the encryption key directly for its secret binding. If you generate it yourself, run the following once in a trusted local terminal and transfer the output privately into that setting. Keep the key out of source control, logs and chat.

   ```sh
   node -e "console.log(require('node:crypto').randomBytes(32).toString('base64'))"
   ```

5. Publish the new Site with its **owner-only** audience and the included D1 migration. Verify that the release succeeded, the `DB` binding exists, both `onedrive_auth` and `drive_transfers` tables were created, and MCP is enabled. The publishing service provisions the private plugin. In **Plugins → Personal → Created by you**, select your new Site's plugin and install or connect it if needed.

The Worker identifies the caller through a header supplied by the Sites authentication gateway. A different host needs an authenticated gateway and origin isolation before accepting requests. Directly exposing this Worker to the internet would let clients supply that header themselves.

## Connect your Microsoft account

1. Open **App registrations → New registration**. Choose **Personal Microsoft accounts only** and leave the redirect URI blank. [Microsoft Entra admin center](https://entra.microsoft.com/)
2. Under **Authentication**, enable **Allow public client flows** and save.
3. Under **API permissions**, add **Microsoft Graph → Delegated permissions → Files.ReadWrite**. Copy the **Application (client) ID** from **Overview**. The app requests `Files.ReadWrite` and `offline_access` at sign-in so it can work with your drive and renew access. Device-code sign-in uses the application ID without a client secret. [Delegated permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference#filesreadwrite)
4. Open your private Site, enter the application ID, and select **Get sign-in code**. Select **Open Microsoft sign-in**, enter the displayed code and approve access to your personal account. Leave the private Site open until it displays **Connected to OneDrive**.
5. In a chat, select the installed plugin and ask **“List the top-level folders in my personal OneDrive.”** This tests the live account and returns the actual root folder listing. For an optional write check, create a uniquely named temporary text file, read back its contents in a second request and ask to move just that file to the recycle bin.

The current tools address the signed-in user's `/me/drive`. Adding `Files.ReadWrite.All` in Entra does not change the scopes this plugin requests or add tools to discover other people's shared drives. Use **Disconnect Microsoft account** on your private Site to remove the stored local connection. You can also manage the app's consent in your Microsoft account.

## Source and distribution

The `overlay/` directory contains the deployed application's code and tests. `apply-overlay.mjs` applies them to a fresh compatible starter. You can keep this source package in version control independently of your running Site. Publish only the package contents when sharing source. The managed deployment checkout also contains a project identifier and can accumulate local credentials and build or database state.

The application source, installer, documentation and tests in this repository are available under the [MIT License](LICENSE). Obtain the Sites starter separately through Sites and retain its supplied notices.
