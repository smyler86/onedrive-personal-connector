import assert from "node:assert/strict";
import test from "node:test";
import { attachmentTools, callAttachment } from "../lib/attachment.mjs";

const file = { file_id: "file_01", download_url: "https://files.oaiusercontent.com/attachment", file_name: "photo.jpg", mime_type: "image/jpeg" };
const neverGraph = () => { throw new Error("Unexpected OneDrive call"); };
const stream = (...chunks) => new ReadableStream({start(controller) { for (const chunk of chunks) controller.enqueue(chunk); controller.close(); }});
const graphItem = () => Response.json({id:"item!1",name:"photo.jpg",eTag:'"fresh"',file:{mimeType:"image/jpeg"},size:4});

test("both tools declare a file parameter in the supported ChatGPT schema", () => {
  for (const tool of attachmentTools) {
    assert.deepEqual(tool._meta["openai/fileParams"], ["file"]);
    assert.deepEqual(tool.inputSchema.properties.file.required, ["download_url", "file_id"]);
    assert.deepEqual(Object.keys(tool.inputSchema.properties.file.properties).sort(), ["download_url","file_id","file_name","mime_type"]);
  }
});

test("save an attachment using original bytes and filename; preserve an existing file with that name", async () => {
  let sourceCalls = 0, writes = 0;
  const source = async (url, init) => {
    sourceCalls++;
    assert.equal(new Headers(init.headers).has("Authorization"), false);
    assert.equal(init.redirect, "manual");
    assert.equal(init.credentials, "omit");
    if (sourceCalls === 1) return Response.redirect("https://storage.oaiusercontent.com/file", 307);
    return new Response(stream(Uint8Array.of(0,21), Uint8Array.of(204,0)));
  };
  const result = await callAttachment("save_attachment", {file}, async (path, init) => {
    writes++;
    assert.equal(new URL(path, "https://graph.microsoft.com").searchParams.get("@microsoft.graph.conflictBehavior"), "fail");
    assert.match(path, /\/root:\/photo\.jpg:\/content/);
    assert.deepEqual([...init.body], [0,21,204,0]);
    return graphItem();
  }, source);
  assert.equal(sourceCalls, 2);
  assert.equal(writes, 1);
  assert.equal(JSON.parse(result[0].text).id, "item!1");
});

test("replace requires exact revision and targets the original item", async () => {
  const result = await callAttachment("replace_with_attachment", {file, item_id:"ABC!1",etag:'"version2"'}, async (path, init) => {
    assert.equal(path, "/me/drive/items/ABC!1/content");
    assert.equal(new Headers(init.headers).get("If-Match"), '"version2"');
    return graphItem();
  }, async () => new Response(new Uint8Array(0), {headers:{"Content-Length":"0"}}));
  assert.ok(result[0].text);
  await assert.rejects(callAttachment("replace_with_attachment", {file,item_id:"ABC!1"}, neverGraph, async () => {throw new Error("Unexpected fetch");}), /required/);
});

test("non-OpenAI URLs, including redirects, are rejected before any bytes reach OneDrive", async () => {
  for (const url of ["https://unrelated.example/item","https://oaisdmntpr.wrong.blob.core.windows.net/item",
    "https://somenameaws.s3.us-east-1.amazonaws.com/item","https://oaisdmntprprod.s3.us-east-1.amazonaws.com/item",
    "https://oaisdmntprprodaws.s3.us-east-1.amazonaws.com.unrelated.example/item","http://files.oaiusercontent.com/file",
    "https://a:b@files.oaiusercontent.com/file","https://files.oaiusercontent.com:444/file",
    "https://files.oaiusercontent.com.example/file","https://chatgpt.com/settings",
    "https://unrelated.blob.core.windows.net/file"]) {
    let reads = 0;
    await assert.rejects(callAttachment("save_attachment", {file:{...file,download_url:url}}, neverGraph, async () => {reads++;}), /attachment supplied/);
    assert.equal(reads,0);
  }
  let reads = 0;
  await assert.rejects(callAttachment("save_attachment", {file}, neverGraph, async () => {reads++; return Response.redirect("https://unrelated.example/secret");}), /attachment supplied/);
  assert.equal(reads,1);
});

test("refuses an oversized attachment, oversized stream without length, or incomplete body before writing", async () => {
  const oversized = 20*1024*1024+1;
  for (const make of [
    () => new Response(null,{headers:{"Content-Length":String(oversized)}}),
    () => new Response(stream(new Uint8Array(oversized))),
    () => new Response(Uint8Array.of(1),{headers:{"Content-Length":"4"}}),
  ]) await assert.rejects(callAttachment("save_attachment",{file},neverGraph,make), /20 MiB|incomplete/);
});

test("files missing a filename accept an explicit destination name; graph refuses changed revision",async () => {
  await assert.rejects(callAttachment("save_attachment",{file:{...file,file_name:undefined}},neverGraph), /filename/);
  await assert.rejects(callAttachment("replace_with_attachment",{file,item_id:"ABC!1",etag:'"v"'},async () => new Response(null,{status:412}),async () => new Response("abc")), /changed|412/);
});

test("direct file transfer accepts only the supported OpenAI Azure and S3 storage host patterns", async () => {
  for (const host of ["oaisdmntprwestus.blob.core.windows.net", "oaisdsorpreast1.blob.core.windows.net",
    "oaisdmntpruseast1aws.s3.us-east-1.amazonaws.com"]) {
    let sourceCalls = 0;
    const result = await callAttachment("save_attachment", {file:{...file,download_url:"https://" + host + "/file"}},
      async () => graphItem(), async () => { sourceCalls++; return new Response("ok"); });
    assert.equal(sourceCalls, 1);
    assert.equal(JSON.parse(result[0].text).id, "item!1");
  }
});
