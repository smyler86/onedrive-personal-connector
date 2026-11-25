import { putFileBytes } from "./drive.mjs";

const maxBytes = 20 * 1024 * 1024;
const limitMessage = "Direct attachments support up to 20 MiB. Use open_upload_page for a larger file.";
const fileSchema = {
  type: "object",
  properties: {
    download_url: { type: "string" },
    file_id: { type: "string" },
    mime_type: { type: "string" },
    file_name: { type: "string" },
  },
  required: ["download_url", "file_id"],
  additionalProperties: false,
};

const id = { type: "string", minLength: 1, maxLength: 512 };
const name = { type: "string", minLength: 1, maxLength: 255 };
function descriptor(toolName, description, properties, required, destructive) {
  return {
    name: toolName,
    description,
    inputSchema: { type: "object", additionalProperties: false,
      properties: { file: fileSchema, ...properties }, required: ["file", ...required] },
    annotations: { readOnlyHint: false, destructiveHint: destructive, openWorldHint: true },
    _meta: { "openai/fileParams": ["file"] },
  };
}

export const attachmentTools = [
  descriptor("save_attachment", "Save a selected ChatGPT attachment, up to 20 MiB, as a new OneDrive file. Use the attachment's filename or set name. Omit parent_id for the drive root. A name collision fails.", {
    parent_id: { ...id, description: "Destination folder ID, or root." },
    name: { ...name, description: "Destination filename when overriding the attachment's filename or when it is missing." },
  }, [], false),
  descriptor("replace_with_attachment", "Replace an existing OneDrive file's entire contents with a selected ChatGPT attachment up to 20 MiB. Identify the file by its ID and most recent etag.", {
    item_id: id,
    etag: { ...id, description: "Exact current etag returned by get_item or list_folder." },
  }, ["item_id", "etag"], true),
];

function sourceLocation(value, current) {
  let url;
  try { url = new URL(value, current); } catch { /* Invalid or missing entry. */ }
  const allowed = url && (url.hostname.endsWith(".oaiusercontent.com") ||
    /^(?:oaisdmntpr|oaisdsorpr)[a-z0-9]*\.blob\.core\.windows\.net$/.test(url.hostname) ||
    /^oaisdmntpr[a-z0-9-]*aws\.s3\.[a-z0-9-]+\.amazonaws\.com$/.test(url.hostname) ||
    (url.hostname === "chatgpt.com" && url.pathname === "/backend-api/estuary/content"));
  if (!allowed || url.protocol !== "https:" || url.port || url.username || url.password) {
    throw new Error("Use an attachment supplied by ChatGPT. If its file host is unsupported, use open_upload_page.");
  }
  return url;
}

async function boundedBytes(response) {
  if (response.status !== 200) {
    await response.body?.cancel();
    throw new Error("ChatGPT could not supply this file. Reattach it or use open_upload_page.");
  }
  const declared = response.headers.get("Content-Length");
  const encoded = response.headers.get("Content-Encoding");
  const exact = declared !== null && (!encoded || encoded.toLowerCase() === "identity");
  const knownLength = exact ? Number(declared) : undefined;
  if (exact && (!/^\d+$/.test(declared) || !Number.isSafeInteger(knownLength) || knownLength > maxBytes)) {
    await response.body?.cancel();
    throw new Error(limitMessage);
  }
  // Bound total memory even if the source omits its length or emits many small chunks.
  const bytes = new Uint8Array(knownLength ?? maxBytes);
  let size = 0;
  if (response.body) {
    const reader = response.body.getReader();
    try {
      while (true) {
        const {done, value} = await reader.read().catch(() => {
          throw new Error("ChatGPT could not finish downloading the attachment. Reattach the file.");
        });
        if (done) break;
        if (size + value.length > bytes.length) {
          await reader.cancel().catch(() => {});
          throw new Error(size + value.length > maxBytes ? limitMessage : "The attachment's size changed during transfer. Reattach the file.");
        }
        bytes.set(value, size);
        size += value.length;
      }
    } finally { reader.releaseLock(); }
  }
  if (exact && size !== knownLength) throw new Error("The attachment was incomplete. Reattach the file.");
  return bytes.subarray(0, size);
}

async function fetchAttachment(file, fetcher) {
  if (!file || typeof file.file_id !== "string" || !file.file_id.trim() ||
      file.file_id.length > 1024 || typeof file.download_url !== "string" || !file.download_url) {
    throw new Error("Select the file attachment to save.");
  }
  let url = sourceLocation(file.download_url);
  const signal = AbortSignal.timeout(60000);
  for (let attempt = 0; attempt < 5; attempt++) {
    let response;
    try {
      response = await fetcher(url.toString(), {method: "GET", headers: {"Accept-Encoding": "identity"},
        credentials: "omit", redirect: "manual", signal});
    } catch { throw new Error("ChatGPT could not supply this attachment. Reattach it or use open_upload_page."); }
    if ([301, 302, 303, 307, 308].includes(response.status)) {
      await response.body?.cancel();
      const location = response.headers.get("location");
      if (!location) break;
      url = sourceLocation(location, url);
      continue;
    }
    return boundedBytes(response);
  }
  throw new Error("The attachment could not be downloaded. Use open_upload_page.");
}

export async function callAttachment(toolName, args, graph, fetcher = fetch) {
  const spec = attachmentTools.find(tool => tool.name === toolName);
  if (!spec || !args || typeof args !== "object" || Array.isArray(args) ||
      Object.keys(args).some(key => !Object.hasOwn(spec.inputSchema.properties, key)) ||
      spec.inputSchema.required.some(key => !Object.hasOwn(args, key))) {
    throw new Error("Check the attachment tool and its required arguments.");
  }
  const target = toolName === "save_attachment"
    ? {parent_id: args.parent_id ?? "root", name: args.name ?? args.file?.file_name}
    : {item_id: args.item_id, etag: args.etag};
  if (toolName === "save_attachment" && !target.name) throw new Error("Supply a destination filename for this attachment.");
  const contents = await fetchAttachment(args.file, fetcher);
  return putFileBytes(graph, target, contents);
}
