"use client";

import { useEffect, useRef, useState, type FormEvent } from "react";

type Pending = { userCode: string; verificationUri: string; expiresIn: number; interval: number };
type Status = { connected: boolean; pending: boolean; clientId?: string | null } & Partial<Pending>;
type View = { state: "loading" | "ready" | "pending" | "connected" | "unavailable"; device?: Pending };

export async function jsonRequest<T>(path: string, options: RequestInit = {}): Promise<T> {
  const response = await fetch(path, { cache: "no-store", credentials: "same-origin", ...options });
  let data: unknown;
  try {
    data = await response.json();
  } catch {
    throw new Error("The server could not complete this request. Reload this page and try again.");
  }
  if (!response.ok) {
    const message = data && typeof data === "object" && "error" in data ? data.error : undefined;
    throw new Error(typeof message === "string" ? message : `The request failed (${response.status}).`);
  }
  return data as T;
}

function jsonPost(data: unknown): RequestInit {
  return { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) };
}

function validDevice(value: Partial<Pending>): value is Pending {
  return typeof value.userCode === "string" && value.userCode.length > 0 &&
    typeof value.verificationUri === "string" && Number.isFinite(value.interval) && Number.isFinite(value.expiresIn);
}

function signInLink(uri: string) {
  try {
    const url = new URL(uri);
    if (url.protocol === "https:" && !url.username && !url.password && !url.port &&
      ["microsoft.com", "www.microsoft.com", "login.microsoftonline.com", "login.live.com"].includes(url.hostname)) {
      return url.href;
    }
  } catch { /* Use Microsoft's published device sign-in page. */ }
  return "https://microsoft.com/devicelogin";
}

function errorText(error: unknown) {
  return error instanceof Error ? error.message : "Could not contact the server.";
}

export function useConnection() {
  const [view, setView] = useState<View>({ state: "loading" });
  const [clientId, setClientId] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [reloadVersion, setReloadVersion] = useState(0);
  const actionVersion = useRef(0);
  const pollVersion = useRef(0);

  useEffect(() => {
    const controller = new AbortController();
    jsonRequest<Status>("/api/connect/status", { signal: controller.signal }).then(status => {
      if (controller.signal.aborted) return;
      setClientId(status.clientId || "");
      setView(status.pending && validDevice(status) ? { state: "pending", device: status } :
        { state: status.connected ? "connected" : "ready" });
    }).catch(reason => {
      if (!controller.signal.aborted) {
        setView({ state: "unavailable" });
        setError(errorText(reason));
      }
    });
    return () => controller.abort();
  }, [reloadVersion]);

  useEffect(() => {
    if (view.state !== "pending" || !view.device) return;
    const device = view.device;
    const generation = ++pollVersion.current;
    const deadline = Date.now() + device.expiresIn * 1000;
    const controller = new AbortController();
    let timer: ReturnType<typeof setTimeout>;
    const present = () => !controller.signal.aborted && pollVersion.current === generation;
    const poll = async (interval: number) => {
      if (Date.now() >= deadline) {
        if (present()) {
          setView({ state: "ready" });
          setError("The sign-in code expired. Start again for a new code.");
        }
        return;
      }
      try {
        const result = await jsonRequest<{ status: string; interval?: number }>("/api/connect/poll", {
          ...jsonPost({}), signal: controller.signal,
        });
        if (!present()) return;
        if (result.status === "connected") {
          setError("");
          setView({ state: "connected" });
          return;
        }
        if (result.status !== "pending") throw new Error("Sign-in was interrupted. Start again.");
        timer = setTimeout(() => void poll(result.interval || interval), Math.max(1, result.interval || interval) * 1000);
      } catch (reason) {
        if (!present()) return;
        setView({ state: "unavailable" });
        setError(errorText(reason));
      }
    };
    timer = setTimeout(() => void poll(device.interval), Math.max(1, device.interval) * 1000);
    return () => { clearTimeout(timer); controller.abort(); };
  }, [view]);

  async function start(event: FormEvent) {
    event.preventDefault();
    if (busy) return;
    const version = ++actionVersion.current;
    setBusy(true);
    setError("");
    try {
      const device = await jsonRequest<Pending>("/api/connect/start", jsonPost({ clientId: clientId.trim() }));
      if (!validDevice(device)) throw new Error("Microsoft did not return a usable sign-in code.");
      if (version === actionVersion.current) setView({ state: "pending", device });
    } catch (reason) {
      if (version === actionVersion.current) setError(errorText(reason));
    } finally {
      if (version === actionVersion.current) setBusy(false);
    }
  }

  async function disconnect() {
    if (busy) return;
    ++actionVersion.current;
    ++pollVersion.current;
    setBusy(true);
    setError("");
    try {
      await jsonRequest("/api/connect/disconnect", jsonPost({}));
      setView({ state: "ready" });
    } catch (reason) {
      setView({ state: "unavailable" });
      setError(errorText(reason));
    } finally { setBusy(false); }
  }

  function reload() {
    setError("");
    setView({ state: "loading" });
    setReloadVersion(v => v + 1);
  }

  return { view, clientId, setClientId, busy, error, start, disconnect, reload };
}

export function ConnectionPanel({ connection }: { connection: ReturnType<typeof useConnection> }) {
  const { view, clientId, setClientId, busy, error, start, disconnect, reload } = connection;
  const device = view.device;
  const title = {
    loading: "Checking connection", unavailable: "Connection unavailable", connected: "Connected to OneDrive",
    ready: "Connect your account", pending: "Finish Microsoft sign-in",
  }[view.state];
  return <section className="od-panel" aria-labelledby="connection-title">
    <div className="od-panel-header"><h2 id="connection-title">{title}</h2></div>
    {view.state === "loading" && <p role="status">Checking your Microsoft connection…</p>}
    {view.state === "ready" && <>
      <p>Register a Microsoft application once, then sign in with the personal account that owns your drive.</p>
      <details className="od-steps">
        <summary>Set up your Microsoft application</summary>
        <ol>
          <li>Open <a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noreferrer">Microsoft Entra app registrations ↗</a> and choose <strong>New registration</strong>. Select the account type for <strong>personal Microsoft accounts</strong>. Leave the redirect URI blank.</li>
          <li>Under <strong>Authentication</strong>, set <strong>Allow public client flows</strong> to <strong>Yes</strong> and save.</li>
          <li>Under <strong>API permissions</strong>, add the Microsoft Graph delegated permission <code>Files.ReadWrite</code>.</li>
          <li>Copy the <strong>Application (client) ID</strong> from the application’s Overview page.</li>
        </ol>
      </details>
      <form onSubmit={start}>
        <label htmlFor="client-id">Application (client) ID</label>
        <input id="client-id" type="text" autoComplete="off" spellCheck={false} placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" required pattern="[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}" value={clientId} onChange={e => setClientId(e.target.value)} title="Enter the Microsoft application ID, including hyphens." disabled={busy}/>
        <button className="od-primary" disabled={busy}>{busy ? "Starting sign-in…" : "Get sign-in code"}</button>
      </form>
    </>}
    {view.state === "pending" && device && <>
      <p>Enter this code on Microsoft’s sign-in page and approve access to your OneDrive.</p>
      <output className="od-code" aria-label="Microsoft sign-in code">{device.userCode}</output>
      <div className="od-actions"><a className="od-primary" href={signInLink(device.verificationUri)} target="_blank" rel="noreferrer">Open Microsoft sign-in ↗</a>
        <button type="button" className="od-text-button" onClick={disconnect} disabled={busy}>Cancel sign-in</button></div>
      <p className="od-small" role="status">Waiting for Microsoft. Leave this page open while you sign in.</p>
    </>}
    {view.state === "connected" && <>
      <p>Your personal Microsoft account is connected. The plugin can create, read, update, move, and delete OneDrive files and folders.</p>
      <button type="button" className="od-secondary" onClick={disconnect} disabled={busy}>{busy ? "Disconnecting…" : "Disconnect Microsoft account"}</button>
    </>}
    {error && <p className="od-error" role="alert">{error}</p>}
    {view.state === "unavailable" && <button type="button" className="od-secondary" onClick={reload}>Check connection again</button>}
  </section>;
}
