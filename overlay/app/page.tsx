"use client";

import { ConnectionPanel, useConnection } from "./drive-status";
import { TransferPanel } from "./drive-transfer";

export default function Home() {
  const connection = useConnection();
  const connected = connection.view.state === "connected";
  return <div className="od-shell">
    <header className="od-top">
      <div className="od-top-inner"><span className="od-wordmark">Personal OneDrive</span>
        <a href="https://onedrive.live.com/" target="_blank" rel="noreferrer">Open OneDrive ↗</a>
      </div>
    </header>
    <main className="od-main">
      <div className="od-intro"><h1>{connected ? "Your OneDrive workspace" : "Your OneDrive connection"}</h1>
        <p>{connected ? "Manage the connection or transfer a file from your device." : "Connect your personal drive to ChatGPT to work with your documents, photos, and videos."}</p>
      </div>
      <div className="od-grid"><div className="od-primary-col">
        {connected && <TransferPanel/>}
        <ConnectionPanel connection={connection} />
      </div>
      <aside className="od-aside" aria-labelledby="usage-title"><h2 id="usage-title">Use it in ChatGPT</h2>
        <p>In a chat, type <code>@</code> and select <strong>Personal OneDrive</strong>. You can search your drive, create folders, rename files, and move items to the recycle bin.</p>
        <blockquote>“List the most recently modified files in my Documents folder.”</blockquote>
        <h3>Photos and videos</h3>
        <p>Ask for a preview or download link. For a large upload, ask the plugin for a transfer link to the destination folder.</p>
      </aside></div>
    </main>
  </div>;
}
