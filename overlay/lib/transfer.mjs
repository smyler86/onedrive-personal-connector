/* The OAuth token belongs exclusively to Graph. Revocable OneDrive upload
 * session URLs never leave this worker or appear in its connection logs. */
const CHUNK = 5 * 1024 * 1024;
const SEGMENT = 320 * 1024;
const MAX_FILE = 250 * 1024 ** 3;
const textEncoder = new TextEncoder();
const REDIRECTS = new Set([301, 302, 303, 307, 308]);
const PROVIDER_HOSTS = ["1drv.com", "onedrive.live.com", "storage.live.com", "livefilestore.com",
  "api.onedrive.com", "docs.live.net", "storage.msn.com", "skydrive.live.com"];

class TransferError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

function fail(status, message) {
  throw new TransferError(status, message);
}

function resourceUrl(value) {
  let target;
  try { target = new URL(value); } catch { fail(502, "OneDrive returned an invalid resource URL."); }
  if (target.protocol !== "https:" || target.port || target.username || target.password
    || !PROVIDER_HOSTS.some(host => target.hostname === host || target.hostname.endsWith("." + host)) && target.hostname !== "my.microsoftpersonalcontent.com") {
    fail(502, "OneDrive returned an unsupported transfer endpoint.");
  }
  return target;
}

function validRange(value) {
  if (value && (!/^bytes=(?:\d+-\d*|-\d+)$/.test(value) || value.length > 50)) {
    fail(416, "Use a single byte range.");
  }
  return value;
}

// The source response must be a Graph response or a previously authorised
// provider response. Only the Range header is copied across host boundaries.
export async function followResourceRedirect(response, headers = {}, fetcher = fetch) {
  const requested = validRange(new Headers(headers).get("range"));
  const downloadHeaders = requested ? { Range: requested } : {};
  for (let index = 0; REDIRECTS.has(response.status); index++) {
    if (index === 4) fail(502, "OneDrive returned too many transfer redirects.");
    const next = resourceUrl(response.headers.get("location"));
    await response.body?.cancel();
    response = await fetcher(next, { method: "GET", headers: downloadHeaders, redirect: "manual" });
  }
  return response;
}

function idOf(value) {
  if (typeof value !== "string" || !/^[A-Za-z0-9_!.~-]{1,240}$/.test(value) || value === "." || value === "..") {
    fail(400, "A valid OneDrive item ID is required.");
  }
  return value;
}

function fileName(value) {
  if (typeof value !== "string" || !value.trim() || value.length > 255 || value !== value.trim()
    || value.endsWith(".") || /[\x00-\x1f\x7f"*:<>?\/\\|]/.test(value)) {
    fail(400, "Provide a valid OneDrive file name without a path.");
  }
  return value;
}

function revision(value) {
  if (typeof value !== "string" || !value.length || value.length > 512 || /[\x00-\x1f\x7f]/.test(value)
    || value.includes("*") || value.trim() !== value || value === "") {
    fail(400, "Provide the selected file's current eTag.");
  }
  return value;
}

function compactItem(value) {
  if (typeof value?.id !== "string" || typeof value?.name !== "string") {
    fail(502, "OneDrive did not return an item for the operation.");
  }
  const { id, name, eTag, size } = value;
  return { id, name, eTag, size };
}

function versionedResult(item, size) {
  return { complete: true, offset: size, item: compactItem(item) };
}

function graphError(response, context) {
  const status = response.status;
  const reasons = {
    401: "Microsoft sign-in is required again.", 403: "Microsoft denied permission.",
    404: "The requested item or transfer was not found.", 409: "The destination has a conflicting item or the upload changed.",
    410: "The resource is no longer available.", 412: "The file changed. Read its current version before retrying.",
    416: "The requested range is unavailable. Refresh the active session status before retrying.",
    429: "OneDrive is rate limiting requests. Retry shortly.", 507: "OneDrive has insufficient storage."
  };
  fail(status >= 400 && status <= 599 ? status : 502, context + " " + (reasons[status] || "Microsoft returned HTTP " + status + "."));
}

function checked(response, context) {
  if (!response.ok) graphError(response, context);
  return response;
}

function resourcePath(id) {
  return "/me/drive/items/" + encodeURIComponent(idOf(id));
}

export async function downloadContent(graph, itemId, range, fetcher = fetch) {
  const headers = validRange(range) ? { Range: range } : {};
  const response = await graph(resourcePath(itemId) + "/content", { headers, redirect: "manual" });
  return followResourceRedirect(response, headers, fetcher);
}

async function download(url, request, graph, fetcher) {
  const id = idOf(url.searchParams.get("id"));
  const meta = await checked(await graph(resourcePath(id) + "?$select=id,name,file"), "Cannot read the download.").json();
  if (!meta.file) fail(400, "Select a file to download.");
  const response = await downloadContent(graph, id, request.headers.get("range"), fetcher);
  if (!response.ok) {
    if (response.status === 416) {
      const e = new Headers({ "cache-control": "no-store" });
      const actual = response.headers.get("content-range");
      if (actual) e.set("content-range", actual);
      return new Response(null, { status: 416, headers: e });
    }
    graphError(response, "Cannot download the file.");
  }
  const name = String(meta.name || "download").slice(0, 255);
  const asciiName = name.replace(/[^A-Za-z0-9._ -]/g, "_") || "download";
  const encodedName = encodeURIComponent(name).replace(/['()*]/g, c => "%" + c.charCodeAt(0).toString(16));
  const headers = new Headers({
    "content-type": "application/octet-stream",
    "content-disposition": 'attachment; filename="' + asciiName + '"; filename*=UTF-8\'\'' + encodedName,
    "content-security-policy": "default-src 'none'; sandbox",
    "x-content-type-options": "nosniff", "cache-control": "no-store",
    "referrer-policy": "no-referrer", "accept-ranges": "bytes"
  });
  for (const field of ["content-length", "content-range", "etag", "last-modified"]) {
    if (response.headers.has(field)) headers.set(field, response.headers.get(field));
  }
  return new Response(response.body, { status: response.status, headers });
}

function base64(bytes) {
  return btoa(String.fromCharCode(...bytes));
}

function unbase64(value) {
  return Uint8Array.from(atob(value), c => c.charCodeAt(0));
}

async function encryptionKey(env) {
  let key;
  try { key = unbase64(env.DRIVE_ENCRYPTION_KEY); } catch { /* initialise below */ }
  if (key?.byteLength !== 32) fail(503, "The transfer service is not configured.");
  return crypto.subtle.importKey("raw", key, "AES-GCM", false, ["encrypt","decrypt"]);
}

function envelope(iv, sessionId, owner) {
  return {name: "AES-GCM", iv, additionalData: textEncoder.encode(JSON.stringify(["transfer-v1", sessionId, owner]))};
}

async function storeValue(env, sessionId, owner, value) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encoded = textEncoder.encode(JSON.stringify(value));
  const ciphertext = await crypto.subtle.encrypt(envelope(iv, sessionId, owner), await encryptionKey(env), encoded);
  return base64(iv) + "." + base64(new Uint8Array(ciphertext));
}

async function readValue(env, sessionId, owner, parts) {
  const [iv, ciphertext] = parts.split(".");
  const raw = await crypto.subtle.decrypt(envelope(unbase64(iv), sessionId, owner), await encryptionKey(env), unbase64(ciphertext));
  return JSON.parse(new TextDecoder().decode(raw));
}

function expires(value) {
  const time = Date.parse(value);
  if (!Number.isFinite(time) || time <= Date.now()) fail(502, "OneDrive did not return an active upload session.");
  return time;
}

function internalStatus(data, size) {
  const ranges = data?.nextExpectedRanges;
  if (!Array.isArray(ranges) || !ranges.length) fail(502, "The upload status contains no expected ranges.");
  const parts = ranges.map(range => {
    const match = /^(\d+)-(?:\d+)?$/.exec(range);
    if (!match) fail(502, "OneDrive returned an unexpected upload range.");
    return Number(match[1]);
  });
  const offset = Math.min(...parts);
  if (!Number.isSafeInteger(offset) || offset < 0 || offset > size) fail(502, "The expected upload offset is invalid.");
  return { complete: false, offset };
}

function json(body, status = 200) {
  return Response.json(body, { status, headers: { "cache-control":"no-store", "x-content-type-options":"nosniff" } });
}

async function currentDriveId(graph) {
  const drive = await checked(await graph("/me/drive?$select=id"), "Cannot verify the active OneDrive connection.").json();
  if (typeof drive.id !== "string" || !drive.id.length || drive.id.length > 512) fail(502,"OneDrive did not identify the current drive.");
  return drive.id;
}

async function requestBytes(request, limit, exact = false) {
  const contentLength = request.headers.get("content-length");
  if (contentLength !== null && (!/^\d+$/.test(contentLength) || Number(contentLength) > limit)) {
    fail(413, "The upload exceeds the expected size.");
  }
  const buffer = new Uint8Array(limit);
  let offset = 0;
  if (request.body) {
    const reader = request.body.getReader();
    try {
      while (true) {
        const {value,done} = await reader.read();
        if (done) break;
        if (offset + value.length > limit) {
          await reader.cancel();
          fail(413, "The upload exceeds the expected size.");
        }
        buffer.set(value,offset);
        offset += value.length;
      }
    } finally { reader.releaseLock(); }
  }
  if (exact && offset !== limit) fail(400, "The chunk length does not match its Content-Range.");
  return buffer.subarray(0,offset);
}

async function createUpload(request, env, owner, graph) {
  if (!(request.headers.get("content-type") || "").startsWith("application/json")) fail(415, "Send a JSON upload target.");
  let body;
  const raw = await requestBytes(request, 8192);
  try { body = JSON.parse(new TextDecoder().decode(raw)); } catch { fail(400, "Invalid upload target."); }
  if (!body || typeof body !== "object" || Array.isArray(body)) fail(400, "Invalid upload target.");
  const { size } = body;
  if (!Number.isSafeInteger(size) || size < 0 || size > MAX_FILE) fail(400, "Invalid file size (maximum 250 GiB).");
  const keys = ["size","name","parentId","itemId","eTag"];
  if (Object.keys(body).some(field => !keys.includes(field))) fail(400, "Unrecognised upload target field.");
  const driveId = await currentDriveId(graph);
  let path;
  let name;
  let identity;
  const headers = {};
  if (body.itemId !== undefined) {
    const itemId = idOf(body.itemId);
    const etag = revision(body.eTag);
    if (body.parentId && body.parentId !== "root") fail(400,"Choose a replacement item or a destination folder.");
    const status = await checked(await graph(resourcePath(itemId) + "?$select=id,name,eTag,file"), "Cannot verify the replacement target.").json();
    if (status.id !== itemId || !status.file) fail(400, "The replacement target must be a file.");
    if (status.eTag !== etag) fail(412, "The file changed. Read its current version before replacing it.");
    name = fileName(status.name);
    identity = {itemId, eTag: etag};
    path = resourcePath(itemId);
    headers["If-Match"] = etag;
  } else {
    if (body.eTag !== undefined) fail(400, "A replacement target requires an item ID.");
    name = fileName(body.name);
    const parent = !body.parentId || body.parentId === "root" ? "/me/drive/root" : resourcePath(body.parentId);
    path = parent + ":/" + encodeURIComponent(name) + ":";
  }
  if (size === 0) {
    const suffix = identity ? "" : "?@microsoft.graph.conflictBehavior=fail";
    const response = await graph(path + "/content" + suffix, {method:"PUT", headers: { ...headers, "Content-Type":"application/octet-stream" }, body:new Uint8Array()});
    return json(versionedResult(await checked(response, "Could not create the empty file.").json(), 0));
  }
  const response = await graph(path + "/createUploadSession", {method:"POST", headers:{...headers,"Content-Type":"application/json"},
    body: JSON.stringify({item:{name, fileSize:size, "@microsoft.graph.conflictBehavior":identity ? "replace" : "fail"}})});
  const session = await checked(response, "Could not start the upload.").json();
  const target = resourceUrl(session.uploadUrl).href;
  const expiresAt = expires(session.expirationDateTime);
  const sessionId = crypto.randomUUID();
  const envelope = await storeValue(env,sessionId,owner,{url:target, driveId, size, ...identity});
  await env.DB.prepare("DELETE FROM drive_transfers WHERE expires_at < ?").bind(Date.now()).run();
  await env.DB.prepare("INSERT INTO drive_transfers (id, owner, session_json, expires_at) VALUES (?,?,?,?)").bind(sessionId,owner,envelope,expiresAt).run();
  return json({sessionId,chunkSize:CHUNK,offset:0,complete:false});
}

async function getSession(env, owner, sessionId, graph) {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/.test(sessionId || "")) {
    fail(400,"Invalid upload session.");
  }
  const row = await env.DB.prepare("SELECT session_json, expires_at FROM drive_transfers WHERE id=? AND owner=?").bind(sessionId,owner).first();
  if (!row) fail(404,"The upload session is unavailable.");
  if (row.expires_at <= Date.now()) fail(410,"The upload session expired. Check the file in OneDrive before restarting.");
  const value = await readValue(env,sessionId,owner,row.session_json);
  if (await currentDriveId(graph) !== value.driveId) fail(403,"This upload belongs to a different OneDrive connection.");
  return value;
}

async function sessionRequest(session, init, fetcher) {
  const response = await fetcher(resourceUrl(session.url), {...init, redirect:"manual"});
  if (REDIRECTS.has(response.status)) fail(502, "OneDrive redirected an upload-session request.");
  if (response.status === 404 || response.status === 410) {
    fail(409, "OneDrive closed this upload session. Verify whether the file arrived before restarting.");
  }
  return response;
}

async function saveExpiry(env, sessionId, owner, status) {
  const current = expires(status.expirationDateTime);
  await env.DB.prepare("UPDATE drive_transfers SET expires_at=MAX(expires_at,?) WHERE id=? AND owner=?").bind(current,sessionId,owner).run();
}

async function handleSession(url, request, env, owner, graph, fetcher) {
  const id = url.searchParams.get("sessionId");
  const session = await getSession(env,owner,id,graph);
  if (request.method === "DELETE") {
    if (!session.finished) {
      const response = await sessionRequest(session,{method:"DELETE"},fetcher);
      checked(response,"Could not cancel the upload.");
    }
    await env.DB.prepare("DELETE FROM drive_transfers WHERE id=? AND owner=?").bind(id,owner).run();
    return json({cancelled:!session.finished, complete:!!session.finished});
  }
  if (session.finished) return json(session.finished);
  if (request.method === "GET") {
    const result = await sessionRequest(session,{method:"GET"},fetcher);
    const status = await checked(result,"Could not determine the upload status.").json();
    await saveExpiry(env,id,owner,status);
    return json(internalStatus(status,session.size));
  }
  const range = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(request.headers.get("content-range") || "");
  if (!range) fail(400,"Invalid upload Content-Range.");
  const [start,end,size] = range.slice(1).map(Number);
  const length = end - start + 1;
  if (![start,end,size].every(Number.isSafeInteger) || size !== session.size || end < start || end >= size || length > CHUNK
    || start % SEGMENT !== 0 || (end < size-1 && length % SEGMENT !== 0)) {
    fail(400,"The upload fragment has an invalid range or length.");
  }
  if (end === size-1 && session.itemId) {
    const current = await checked(await graph(resourcePath(session.itemId) + "?$select=id,eTag"),"Cannot verify the current file version.").json();
    if (current.eTag !== session.eTag) fail(412,"The file changed during the upload. Review the new version before replacing it.");
  }
  const bytes = await requestBytes(request,length,true);
  const response = await sessionRequest(session,{method:"PUT",headers:{"Content-Range":"bytes "+start+"-"+end+"/"+size,
    "Content-Type":"application/octet-stream","Content-Length":String(length)},body:bytes},fetcher);
  const result = await checked(response,"Could not upload this fragment.").json();
  if (response.status === 202) {
    await saveExpiry(env,id,owner,result);
    return json(internalStatus(result,size));
  }
  if (!(response.status === 201 || response.status === 200) || end !== size-1) fail(502,"Unexpected upload completion response.");
  if (session.itemId && result.id !== session.itemId) fail(502,"OneDrive acknowledged an unexpected item.");
  const finished = versionedResult(result,size);
  const value = await storeValue(env,id,owner,{driveId:session.driveId,size,finished});
  await env.DB.prepare("UPDATE drive_transfers SET session_json=?, expires_at=? WHERE id=? AND owner=?")
    .bind(value,Date.now()+5*60*1000,id,owner).run();
  return json(finished);
}

export async function handleTransfer(request, env, userId, graph, fetcher = fetch) {
  try {
    if (!userId) fail(401,"Sign in before accessing OneDrive.");
    const url = new URL(request.url);
    if (request.method !== "GET" && request.headers.get("origin") !== url.origin) fail(403,"Send this request from the connection page.");
    if (url.pathname === "/api/files/download" && request.method === "GET") {
      return await download(url,request,graph,fetcher);
    }
    if (url.pathname === "/api/files/upload/start" && request.method === "POST") {
      return await createUpload(request,env,userId,graph);
    }
    const methods = {"/api/files/upload/status":"GET","/api/files/upload/chunk":"PUT","/api/files/upload/session":"DELETE"};
    if (methods[url.pathname] !== request.method) fail(404,"Unknown transfer endpoint.");
    return await handleSession(url,request,env,userId,graph,fetcher);
  } catch (error) {
    if (error instanceof TransferError) return json({error:error.message},error.status);
    const status = Number(error?.status);
    if (status === 401 || status === 403) return json({error:"Sign in to the active Microsoft account and try again."},status);
    return json({error:"The transfer could not be completed. For an active upload, check its status before retrying."},502);
  }
}

const input = (properties,required = []) => ({type:"object",additionalProperties:false,properties,required});
const textField = description => ({type:"string",description});
export const transferTools = [
  {name:"open_upload_page",description:"Open the private browser uploader for videos, photos, documents or other file types. The user selects the local file and initiates the upload. Supply itemId and eTag together to replace an existing file.",inputSchema:input({
    parentId:textField("A destination folder's ID or root for a new file."),itemId:textField("ID of an existing file to replace."),
    eTag:textField("Current eTag of the file being replaced."),name:textField("Optional original file name shown on the upload page.")}),
    annotations:{readOnlyHint:true,destructiveHint:false,idempotentHint:true,openWorldHint:false}},
  {name:"get_download_link",description:"Return a private link to download a file, including large videos and photos. The link requires the signed-in user's active Microsoft connection.",inputSchema:input({itemId:textField("Microsoft OneDrive file ID.")},["itemId"]),
    annotations:{readOnlyHint:true,destructiveHint:false,idempotentHint:true,openWorldHint:false}}
];

export async function callTransferTool(name,args,origin) {
  const url = new URL("/", origin);
  if (name === "get_download_link") {
    url.pathname = "/api/files/download";
    url.searchParams.set("id",idOf(args.itemId));
  } else if (name === "open_upload_page") {
    if (args.itemId !== undefined || args.eTag !== undefined) {
      if (args.parentId) fail(400,"Choose an upload folder or a replacement item.");
      url.searchParams.set("itemId",idOf(args.itemId));
      url.searchParams.set("eTag",revision(args.eTag));
      if (args.name) url.searchParams.set("name",fileName(args.name));
    } else {
      url.searchParams.set("parentId",idOf(args.parentId || "root"));
    }
  } else fail(400,"Unknown transfer tool.");
  return [{type:"text",text:JSON.stringify({url:url.href, purpose:name === "get_download_link" ? "Download file" : "Upload file"})}];
}
