// Microsoft Graph v1.0 DriveItems, scoped to the signed-in user's default drive.
const base = "/me/drive";
const inlineBytes = 1024 * 1024;
const base64Length = 4 * Math.ceil(inlineBytes / 3);
const fields = "id,name,eTag,size,webUrl,file,folder,package,photo,image,video,createdDateTime,lastModifiedDateTime,parentReference";
const fieldsQuery = "?$select=" + fields;
const idSchema = { type: "string", minLength: 1, maxLength: 512, description: "Exact OneDrive item ID from a search, folder listing or metadata result." };
const parentSchema = { ...idSchema, description: "Destination folder ID. Use root for the top level of the drive." };
const nameSchema = { type: "string", minLength: 1, maxLength: 255, description: "One file or folder name, with its extension where applicable." };
const etagSchema = { type: "string", minLength: 1, maxLength: 512, description: "Copy the exact etag returned by the most recent listing or metadata read of this item." };
const contentSchema = {
  content: { type: "string", maxLength: base64Length, description: "Complete UTF-8 text or base64-encoded original bytes, up to 1 MiB after decoding." },
  encoding: { type: "string", enum: ["text", "base64"], default: "text" },
};
const pagination = {
  limit: { type: "integer", minimum: 1, maximum: 100, default: 30 },
  cursor: { type: "string", minLength: 1, maxLength: 20000, description: "Use the next_cursor from the previous page with the same folder or search." },
};
const formats = ["metadata", "text", "thumbnail", "image", "file", "pdf"];

function tool(name, description, properties, required, readOnly = false, destructive = false) {
  return {
    name, description,
    inputSchema: { type: "object", additionalProperties: false, properties, required },
    annotations: { readOnlyHint: readOnly, destructiveHint: destructive, openWorldHint: true },
  };
}

export const driveTools = [
  tool("list_folder", "List one page of a folder in personal OneDrive. Start at root, then use folder IDs and next_cursor to continue.", { folder_id: parentSchema, ...pagination }, [], true),
  tool("search_files", "Search personal OneDrive by filename, metadata or indexed text. Results include files and folders. Use next_cursor for additional pages.", { query: { type: "string", minLength: 1, maxLength: 512 }, ...pagination }, ["query"], true),
  tool("get_item", "Get an item's metadata, UTF-8 text, original image, thumbnail (including video posters), original binary file, or a PDF rendition of a supported Office file. Inline results are limited to 1 MiB. For larger originals use get_download_link. A thumbnail is a single image.", { item_id: idSchema, format: { type: "string", enum: formats, default: "metadata" } }, ["item_id"], true),
  tool("create_folder", "Create one folder under an existing folder. A name collision fails and preserves the existing item.", { parent_id: parentSchema, name: nameSchema }, ["parent_id", "name"]),
  tool("create_file", "Create a new file up to 1 MiB. Supply text or base64-encoded bytes. A name collision fails. For larger files use open_upload_page.", { parent_id: parentSchema, name: nameSchema, ...contentSchema }, ["parent_id", "name", "content"]),
  tool("replace_file", "Replace an existing file's entire contents, up to 1 MiB. Select the intended file by ID and supply its current etag. For larger replacements use open_upload_page.", { item_id: idSchema, etag: etagSchema, ...contentSchema }, ["item_id", "etag", "content"], false, true),
  tool("update_item", "Rename or move a file or folder within this personal drive. Provide a new name, new parent, or both. Use its current etag.", { item_id: idSchema, etag: etagSchema, name: nameSchema, parent_id: parentSchema }, ["item_id", "etag"], false, true),
  tool("delete_item", "Move the exact file or folder to the OneDrive recycle bin. Deleting a folder also recycles its contents. Supply the intended item's current etag.", { item_id: idSchema, etag: etagSchema }, ["item_id", "etag"], false, true),
];

const byName = new Map(driveTools.map(entry => [entry.name, entry]));
const encoder = new TextEncoder();

function route(id, allowRoot = false) {
  if (id === "root" && allowRoot) return base + "/root";
  if (typeof id !== "string" || id === "root" || !/^[A-Za-z0-9_!.-]{1,512}$/.test(id) || id === "." || id === "..") {
    throw new Error("Provide a valid OneDrive item ID from a listing or search result.");
  }
  return base + "/items/" + encodeURIComponent(id);
}

function checkedName(value) {
  if (typeof value !== "string" || !value.trim() || value.length > 255 || /[\u0000-\u001f<>:"/\\|?*]/u.test(value) || value === "." || value === "..") {
    throw new Error("Specify one valid OneDrive file or folder name.");
  }
  return value;
}

function version(value) {
  if (typeof value !== "string" || !value.trim() || value.length > 512 || /[\r\n\u0000]/u.test(value) || value.trim() === "*") {
    throw new Error("Specify the item's exact current etag, including quotes if present.");
  }
  return value;
}

function asText(value) {
  return [{ type: "text", text: JSON.stringify(value) }];
}

function binTo64(bytes) {
  let result = "";
  for (let i = 0; i < bytes.length; i += 8192) {
    result += String.fromCharCode(...bytes.subarray(i, i + 8192));
  }
  return btoa(result);
}

function bodyBytes({ content, encoding = "text" }) {
  if (typeof content !== "string" || content.length > base64Length) throw new Error("Supply at most 1 MiB of file contents.");
  let bytes;
  if (encoding === "text") {
    bytes = encoder.encode(content);
  } else if (encoding === "base64") {
    if (!/^(?:[A-Za-z0-9+/]*={0,2})$/.test(content) || content.length % 4 !== 0) throw new Error("Provide standard padded base64 file contents.");
    try {
      bytes = Uint8Array.from(atob(content), c => c.charCodeAt(0));
    } catch {
      throw new Error("Provide valid base64 file contents.");
    }
  } else {
    throw new Error("Select text or base64 encoding.");
  }
  if (bytes.length > inlineBytes) throw new Error("Inline uploads are limited to 1 MiB. Use open_upload_page.");
  return bytes;
}

function openLink(item) {
  // Only durable Microsoft browser links belong in tool results. Graph also returns bearer download URLs.
  try {
    const url = new URL(item.webUrl);
    if (url.protocol === "https:" && !url.username && !url.password &&
        /^(?:onedrive\.live\.com|onedrive\.com|[a-z0-9.-]+\.sharepoint\.com)$/i.test(url.hostname) &&
        ![...url.searchParams.keys()].some(k => /authkey|token|signature|^sig$|tempauth/i.test(k))) return url.href;
  } catch { /* Some items have no web URL. */ }
  if (typeof item.id === "string") return "https://onedrive.live.com/?id=" + encodeURIComponent(item.id);
}

function metadata(item) {
  return {
    id: item.id,
    name: item.name,
    type: item.folder ? "folder" : item.package ? "package" : "file",
    etag: item.eTag,
    size: item.size,
    mime_type: item.file?.mimeType,
    child_count: item.folder?.childCount,
    created: item.createdDateTime,
    modified: item.lastModifiedDateTime,
    parent_id: item.parentReference?.id,
    parent_path: item.parentReference?.path,
    taken: item.photo?.takenDateTime,
    width: item.image?.width ?? item.video?.width,
    height: item.image?.height ?? item.video?.height,
    duration_ms: item.video?.duration,
    web_url: openLink(item),
  };
}

async function checked(response) {
  if (response.ok) return response;
  const instructions = {
    401: "Reconnect your Microsoft account.",
    403: "The signed-in account lacks access or OneDrive requires additional authorization.",
    404: "The item or requested rendition was not found.",
    409: "The change conflicts with an existing item or its current state. Inspect the destination.",
    412: "The item changed. Read its metadata and review the new version before retrying.",
    423: "OneDrive reports the item is locked.",
    429: "Microsoft is rate limiting requests. Retry later.",
  };
  await response.body?.cancel().catch(() => {});
  throw new Error("OneDrive request failed (HTTP " + response.status + "). " + (instructions[response.status] ?? "Check the requested operation and item."));
}

async function json(graph, path, init) {
  return (await checked(await graph(path, init))).json();
}

function nextToken(link) {
  if (!link) return null;
  // Reuse only Microsoft's opaque paging parameter. Never let a cursor select a host, path or action.
  const url = new URL(link);
  if (url.origin !== "https://graph.microsoft.com") throw new Error("OneDrive returned an unexpected pagination URL.");
  const token = [...url.searchParams].find(([key]) => /^\$?skiptoken$/i.test(key))?.[1];
  if (!token || token.length > 20000) throw new Error("OneDrive returned an unsupported pagination token.");
  return token;
}

async function list(graph, path, { limit = 30, cursor }) {
  if (!Number.isInteger(limit) || limit < 1 || limit > 100) throw new Error("Page size must be between 1 and 100.");
  const query = new URLSearchParams({ "$select": fields, "$top": String(limit) });
  if (cursor !== undefined) {
    if (typeof cursor !== "string" || !cursor || cursor.length > 20000) throw new Error("Provide a valid pagination cursor.");
    query.set("$skiptoken", cursor);
  }
  const page = await json(graph, path + "?" + query);
  return asText({ items: page.value.map(metadata), next_cursor: nextToken(page["@odata.nextLink"]) });
}

async function limitedContent(response) {
  if (Number(response.headers.get("content-length")) > inlineBytes) {
    await response.body?.cancel();
    throw new Error("The content exceeds the 1 MiB inline limit. Use get_download_link.");
  }
  const chunks = [];
  let size = 0;
  const reader = response.body?.getReader();
  if (reader) {
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        size += value.length;
        if (size > inlineBytes) {
          await reader.cancel();
          throw new Error("The content exceeds the 1 MiB inline limit. Use get_download_link.");
        }
        chunks.push(value);
      }
    } finally { reader.releaseLock(); }
  }
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
  return bytes;
}

async function getItem(graph, { item_id, format = "metadata" }) {
  if (!formats.includes(format)) throw new Error("Select a supported item format.");
  const path = route(item_id, true);
  const item = metadata(await json(graph, path + fieldsQuery));
  if (format === "metadata") return asText(item);
  if (format !== "thumbnail" && item.type !== "file") throw new Error("Choose a file to read its contents.");
  if (!["thumbnail", "pdf"].includes(format) && item.size > inlineBytes) throw new Error("The original file exceeds 1 MiB. Use get_download_link or request its thumbnail.");
  const suffix = format === "thumbnail" ? "/thumbnails/0/large/content" : format === "pdf" ? "/content?format=pdf" : "/content";
  const response = await checked(await graph(path + suffix));
  const bytes = await limitedContent(response);
  const contentType = (response.headers.get("content-type") ?? "").split(";")[0].toLowerCase();
  const mime = format === "pdf" ? "application/pdf" : contentType && contentType !== "application/octet-stream" ? contentType : (item.mime_type ?? "application/octet-stream");
  if (format === "text") {
    try {
      return [...asText(item), { type: "text", text: new TextDecoder("utf-8", { fatal: true }).decode(bytes) }];
    } catch { throw new Error("This file is not UTF-8 text. Request the file or a supported PDF rendition."); }
  }
  if (format === "thumbnail" || format === "image") {
    const signature = imageType(bytes);
    if (!signature) throw new Error("A supported image was not returned. Request the file or its thumbnail.");
    return [...asText({ ...item, rendition: format }), { type: "image", mimeType: signature, data: binTo64(bytes) }];
  }
  return [...asText(item), { type: "resource", resource: {
    uri: "onedrive://items/" + encodeURIComponent(item.id) + (format === "pdf" ? "?format=pdf" : ""),
    mimeType: mime, blob: binTo64(bytes),
  } }];
}

function imageType(bytes) {
  if (bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255) return "image/jpeg";
  if (bytes[0] === 137 && bytes[1] === 80 && bytes[2] === 78 && bytes[3] === 71) return "image/png";
  const prefix = String.fromCharCode(...bytes.subarray(0, 12));
  if (prefix.startsWith("GIF87a") || prefix.startsWith("GIF89a")) return "image/gif";
  if (prefix.startsWith("RIFF") && prefix.slice(8, 12) === "WEBP") return "image/webp";
}

export async function putFileBytes(graph, args, bytes) {
  if (args.etag !== undefined && !("item_id" in args)) throw new Error("Specify an existing item ID when sending an eTag.");
  let path;
  const headers = { "Content-Type": "application/octet-stream" };
  if ("item_id" in args) {
    if ("parent_id" in args || "name" in args) throw new Error("Specify either a target folder and name, or the existing item to replace.");
    path = route(args.item_id) + "/content";
    headers["If-Match"] = version(args.etag);
  } else {
    path = route(args.parent_id, true) + ":/" + encodeURIComponent(checkedName(args.name)) + ":/content?@microsoft.graph.conflictBehavior=fail";
  }
  const response = await json(graph, path, { method: "PUT", headers, body: bytes });
  return asText(metadata(response));
}

export async function callDrive(name, args, graph) {
  const spec = byName.get(name);
  if (!spec) throw new Error("Unknown OneDrive tool.");
  if (!args || typeof args !== "object" || Array.isArray(args)) throw new Error("Provide tool arguments as an object.");
  if (Object.keys(args).some(k => !Object.hasOwn(spec.inputSchema.properties, k)) ||
      spec.inputSchema.required.some(k => !Object.hasOwn(args, k))) throw new Error("Check the required arguments for " + name + ".");
  switch (name) {
    case "list_folder":
      return list(graph, route(args.folder_id ?? "root", true) + "/children", args);
    case "search_files": {
      if (typeof args.query !== "string" || !args.query.trim() || args.query.length > 512) throw new Error("Supply a search term between 1 and 512 characters.");
      const term = encodeURIComponent(args.query.replaceAll("'", "''")).replaceAll("'", "%27");
      return list(graph, base + "/root/search(q='" + term + "')", args);
    }
    case "get_item":
      return getItem(graph, args);
    case "create_folder": {
      const body = { name: checkedName(args.name), folder: {}, "@microsoft.graph.conflictBehavior": "fail" };
      const item = await json(graph, route(args.parent_id, true) + "/children", {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
      });
      return asText(metadata(item));
    }
    case "create_file":
    case "replace_file":
      return putFileBytes(graph, args, bodyBytes(args));
    case "update_item": {
      const path = route(args.item_id);
      const etag = version(args.etag);
      const body = {};
      if (args.name !== undefined) body.name = checkedName(args.name);
      if (args.parent_id !== undefined) {
        const parentPath = route(args.parent_id, true);
        const parent = args.parent_id === "root" ? (await json(graph, parentPath + "?$select=id")).id : args.parent_id;
        if (!parent || parent === args.item_id) throw new Error("Specify a different destination folder.");
        body.parentReference = { id: parent };
      }
      if (!Object.keys(body).length) throw new Error("Supply a new name or destination folder.");
      const item = await json(graph, path, { method: "PATCH", headers: { "If-Match": etag, "Content-Type": "application/json" }, body: JSON.stringify(body) });
      return asText(metadata(item));
    }
    case "delete_item": {
      await checked(await graph(route(args.item_id), { method: "DELETE", headers: { "If-Match": version(args.etag) } }));
      return asText({ id: args.item_id, recycled: true });
    }
  }
}
