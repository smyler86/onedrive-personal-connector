"use client";

import { useEffect, useRef, useState, type FormEvent } from "react";
import { jsonRequest } from "./drive-status";

type TransferTarget = { parentId: string; itemId: string; eTag: string; name: string };
type Item = { name?: string; id?: string; webUrl?: string; web_url?: string };
type Result = { sessionId?: string; chunkSize?: number; offset: number; complete: boolean; item?: Item };
type Session = { sessionId: string; file: File; chunkSize: number; offset: number };
type Phase = "idle" | "starting" | "running" | "pausing" | "paused" | "canceling" | "done";
const emptyTarget = { parentId: "root", itemId: "", eTag: "", name: "" };

function progressLabel(uploaded: number, total: number) {
  if (total === 0) return "0 bytes";
  const format = (bytes: number) => bytes < 1024 ? `${bytes} B` :
    bytes < 1048576 ? `${(bytes / 1024).toFixed(1)} KB` : `${(bytes / 1048576).toFixed(1)} MB`;
  return `${format(uploaded)} of ${format(total)}`;
}

function eMessage(reason: unknown) {
  return reason instanceof Error ? reason.message : "The transfer could not be completed.";
}

function itemLink(item?: Item) {
  const raw = item?.webUrl || item?.web_url;
  if (!raw) return undefined;
  try {
    const url = new URL(raw);
    return url.protocol === "https:" && !url.username && !url.password ? url.href : undefined;
  } catch { return undefined; }
}

function checkOffset(value: number, total: number) {
  if (!Number.isSafeInteger(value) || value < 0 || value > total) {
    throw new Error("OneDrive returned an invalid upload position. Check the transfer before continuing.");
  }
}

export function TransferPanel() {
  const [target, setTarget] = useState<TransferTarget>(emptyTarget);
  const [file, setFile] = useState<File | null>(null);
  const [phase, setPhase] = useState<Phase>("idle");
  const [error, setError] = useState("");
  const [offset, setOffset] = useState(0);
  const [result, setResult] = useState<Item | undefined>();
  const [pausedManually, setPausedManually] = useState(false);
  const [hasTarget, setHasTarget] = useState(false);
  const [showFolderId, setShowFolderId] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const session = useRef<Session | null>(null);
  const controller = useRef<AbortController | undefined>(undefined);
  const pause = useRef(false);
  const lock = useRef(false);
  const holding = phase !== "idle" && phase !== "done";
  const replace = Boolean(target.itemId);

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const selected = {
      itemId: params.get("itemId") || "", eTag: params.get("eTag") || "",
      name: params.get("name") || "", parentId: params.get("parentId") || "root",
    };
    setTarget(selected);
    setHasTarget(Boolean(selected.itemId || selected.parentId !== "root"));
    return () => controller.current?.abort();
  }, []);

  function confirmed(result: Result, selected: File) {
    setOffset(selected.size);
    setPhase("done");
    setResult(result.item);
    setError("");
    session.current = null;
  }

  async function sendChunks(state: Session, signal: AbortSignal) {
    while (state.offset < state.file.size) {
      if (pause.current) { setPhase("paused"); return; }
      const start = state.offset;
      const end = Math.min(state.file.size, start + state.chunkSize);
      const part = await jsonRequest<Result>(`/api/files/upload/chunk?sessionId=${encodeURIComponent(state.sessionId)}`, {
        method: "PUT", signal, headers: { "Content-Range": `bytes ${start}-${end - 1}/${state.file.size}` },
        body: state.file.slice(start, end),
      });
      if (part.complete) { confirmed(part, state.file); return; }
      checkOffset(part.offset, state.file.size);
      if (part.offset <= start) throw new Error("OneDrive did not advance the upload. Check its status before continuing.");
      state.offset = part.offset;
      setOffset(part.offset);
    }
    throw new Error("OneDrive has not confirmed the final file. Check its status before continuing.");
  }

  async function perform(resume: boolean) {
    if (lock.current || !file) return;
    lock.current = true;
    setError("");
    setPausedManually(false);
    pause.current = false;
    const current = new AbortController();
    controller.current = current;
    const signal = current.signal;
    try {
      if (resume && session.current) {
        setPhase("running");
        const currentSession = session.current;
        const status = await jsonRequest<Result>(`/api/files/upload/status?sessionId=${encodeURIComponent(currentSession.sessionId)}`, { signal });
        if (status.complete) { confirmed(status, file); return; }
        checkOffset(status.offset, file.size);
        currentSession.offset = status.offset;
        setOffset(status.offset);
        await sendChunks(currentSession, signal);
      } else {
        if (replace && !target.eTag) throw new Error("Open a fresh replacement link from ChatGPT to identify the original file and version.");
        setPhase("starting");
        setOffset(0);
        setResult(undefined);
        const started = await jsonRequest<Result>("/api/files/upload/start", {
          method: "POST", headers: { "Content-Type": "application/json" }, signal,
          body: JSON.stringify(replace ?
            { itemId: target.itemId, eTag: target.eTag, ...(target.name ? { name: target.name } : {}), size: file.size } :
            { parentId: target.parentId.trim() || "root", name: file.name, size: file.size }),
        });
        if (started.complete) { confirmed(started, file); return; }
        if (!started.sessionId || !started.chunkSize || !Number.isSafeInteger(started.chunkSize) || started.chunkSize <= 0) throw new Error("The server did not provide a usable upload session.");
        checkOffset(started.offset, file.size);
        session.current = { sessionId: started.sessionId, file, chunkSize: started.chunkSize, offset: started.offset };
        setPhase("running");
        setOffset(started.offset);
        await sendChunks(session.current, signal);
      }
    } catch (reason) {
      if (!signal.aborted) {
        setError(eMessage(reason));
        setPhase(session.current ? "paused" : "idle");
      }
    } finally { lock.current = false; }
  }

  function submit(event: FormEvent) {
    event.preventDefault();
    void perform(false);
  }

  async function cancel() {
    if (lock.current || !session.current) return;
    lock.current = true;
    const current = new AbortController();
    controller.current = current;
    setPhase("canceling");
    try {
      const cancellation = await jsonRequest<{ complete: boolean }>(`/api/files/upload/session?sessionId=${encodeURIComponent(session.current.sessionId)}`, { method: "DELETE", signal: current.signal });
      if (cancellation.complete && file) { confirmed({ complete: true, offset: file.size }, file); return; }
      session.current = null;
      setPhase("idle");
      setOffset(0);
      setError("");
      setFile(null);
      if (input.current) input.current.value = "";
    } catch (reason) {
      if (!current.signal.aborted) {
        setError(eMessage(reason));
        setPhase("paused");
      }
    } finally { lock.current = false; }
  }

  function resetCompleted() {
    setFile(null);
    setResult(undefined);
    setPhase("idle");
    setOffset(0);
    if (input.current) input.current.value = "";
  }

  return <section className="od-panel" aria-labelledby="upload-title">
    <div className="od-panel-header"><h2 id="upload-title">{replace ? "Replace a file" : "Upload a file"}</h2></div>
    {phase === "done" ? <div>
      <p role="status"><strong>{result?.name || target.name || file?.name}</strong> was saved to OneDrive.</p>
      {itemLink(result) && <a className="od-secondary" href={itemLink(result)} target="_blank" rel="noreferrer">Open in OneDrive ↗</a>}
      {!replace && <button type="button" className="od-text-button" onClick={resetCompleted}>Upload another file</button>}
    </div> : <>
      {replace && <p>Replacing <strong>{target.name || target.itemId}</strong> will change its contents and keep the original file name.</p>}
      <form onSubmit={submit}>
        {!replace && <>
          <div className="od-destination"><div><span>Destination</span><b>{target.parentId === "root" || !target.parentId ? "My files" : hasTarget ? "Selected OneDrive folder" : "OneDrive folder"}</b></div>
            <button className="od-text-button" type="button" aria-expanded={showFolderId} onClick={() => setShowFolderId(v => !v)} disabled={holding}>{showFolderId ? "Hide folder ID" : "Change folder"}</button></div>
          {showFolderId && <div className="od-field"><label htmlFor="parent-id">OneDrive folder ID</label><input id="parent-id" value={target.parentId} placeholder="root" onChange={e => setTarget({ ...target, parentId: e.target.value })} disabled={holding}/><small>Use <code>root</code> for My files, or paste a folder ID from the plugin.</small></div>}
        </>}
        <label htmlFor="upload-file">{replace ? "Replacement file" : "File"}</label>
        <input ref={input} id="upload-file" className="od-file" type="file" required disabled={holding} onChange={e => { setFile(e.target.files?.[0] || null); setError(""); }} />
        {phase === "idle" && <button disabled={!file || (replace && !target.eTag)} className="od-primary">{replace ? "Replace original file" : "Upload to OneDrive"}</button>}
      </form>
      {replace && !target.eTag && <p className="od-error" role="alert">Open a fresh replacement link from the plugin to identify the current file version.</p>}
      {holding && <>
        <progress className="od-progress" max={file?.size || 1} value={offset} aria-label="File upload progress" />
        <p className="od-small" role="status">{phase === "starting" ? "Preparing upload…" : phase === "canceling" ? "Canceling upload…" : phase === "pausing" ? "Pausing after the current part…" : phase === "paused" ? pausedManually ? "Upload paused." : "Transfer interrupted." : "Uploading…"} {file && progressLabel(offset, file.size)}</p>
        {phase === "running" && <button type="button" className="od-secondary" onClick={() => { pause.current = true; setPausedManually(true); setPhase("pausing"); }}>Pause upload</button>}
        {phase === "paused" && <div className="od-actions"><button type="button" className="od-primary" onClick={() => void perform(true)}>Check and resume</button><button type="button" className="od-text-button" onClick={cancel}>Cancel upload</button></div>}
      </>}
      {error && <p className="od-error" role="alert">{error}</p>}
    </>}
  </section>;
}
