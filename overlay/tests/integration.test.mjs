import assert from 'node:assert/strict';
import test from 'node:test';
import { handleMcp } from '../lib/mcp.mjs';

const origin = 'https://private.example.test';
const user = 'synthetic-site-user';
const requestId = 'test-request-1';

function request(method, params = {}) {
  return new Request(`${origin}/mcp`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ jsonrpc: '2.0', id: requestId, method, params }),
  });
}

function forbidGraph() {
  assert.fail('A request reached Graph during a local-only operation');
}

async function rpc(method, params, graph = forbidGraph, identity = user) {
  const response = await handleMcp(request(method, params), {}, identity, graph);
  assert.equal(response.ok, true);
  const envelope = await response.json();
  assert.equal(envelope.jsonrpc, '2.0');
  assert.equal(envelope.id, requestId);
  return envelope;
}

test('the Site plugin advertises its tool protocol without Microsoft sign-in', async () => {
  const { result } = await rpc('server/discover', {});
  assert.ok(result.supportedVersions.includes('2026-07-28'));
  assert.ok(Object.hasOwn(result.capabilities, 'tools'));
});

test('tool discovery distinguishes file reads, file writes and deletion', async () => {
  const { result } = await rpc('tools/list', {});
  const registry = new Map(result.tools.map((tool) => [tool.name, tool]));
  assert.equal(registry.size, result.tools.length);
  for (const name of ['get_item', 'list_folder', 'search_files', 'open_connection_settings', 'open_upload_page', 'get_download_link']) {
    assert.equal(registry.get(name)?.annotations?.readOnlyHint, true, name);
  }
  for (const name of ['create_folder', 'create_file', 'replace_file', 'update_item', 'delete_item', 'save_attachment', 'replace_with_attachment']) {
    const tool = registry.get(name);
    assert.ok(tool?.inputSchema, name);
    assert.notEqual(tool.annotations?.readOnlyHint, true, name);
  }
  assert.equal(registry.get('delete_item').annotations?.destructiveHint, true);
});

test('connection instructions are available before a Microsoft account is connected', async () => {
  const { result } = await rpc('tools/call', { name: 'open_connection_settings', arguments: {} });
  assert.notEqual(result.isError, true);
  assert.match(JSON.stringify(result.content), /https:\/\//);
});

test('a missing Site identity cannot access personal OneDrive', async () => {
  let graphCalls = 0;
  const graph = async () => {
    graphCalls++;
    return Response.json({ value: [] });
  };
  const response = await handleMcp(request('tools/call', { name: 'list_folder', arguments: {} }), {}, null, graph);
  if (response.ok) {
    const envelope = await response.json();
    assert.ok(envelope.error || envelope.result?.isError, 'an unauthenticated tool call must fail');
  } else {
    assert.ok([401, 403].includes(response.status));
  }
  assert.equal(graphCalls, 0);
});

test('listing the root delegates to the signed-in drive and returns item information', async () => {
  const calls = [];
  const mockGraph = async (path, init) => {
    calls.push({ path, method: init?.method || 'GET' });
    return Response.json({ value: [{ id: 'driveItem-001', name: 'Test photo.jpg', file: { mimeType: 'image/jpeg' }, size: 56, eTag: '"revision1"' }] });
  };
  const { result } = await rpc('tools/call', { name: 'list_folder', arguments: {} }, mockGraph);
  assert.notEqual(result.isError, true);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].method, 'GET');
  assert.equal(calls[0].path.split('?')[0], '/me/drive/root/children');
  assert.match(JSON.stringify(result.content), /Test photo\.jpg/);
});

test('chat tool dispatch preserves file identity and revision across a write cycle', async () => {
  const fileId = 'f-100';
  const folderId = 'p-200';
  const file = { id: fileId, name: 'draft.txt', file: { mimeType: 'text/plain' }, eTag: '"v1"' };
  const operations = [
    { name: 'create_file', args: { parent_id: 'root', name: file.name, content: 'first version' }, method: 'PUT', path: '/me/drive/root:/draft.txt:/content', response: () => Response.json(file, {status: 201}) },
    { name: 'replace_file', args: { item_id: fileId, etag: '"v1"', content: 'second version' }, method: 'PUT', path: `/me/drive/items/${fileId}/content`, rev: '"v1"', response: () => Response.json({ ...file, eTag: '"v2"' }) },
    { name: 'update_item', args: { item_id: fileId, etag: '"v2"', parent_id: folderId, name: 'final.txt' }, method: 'PATCH', path: `/me/drive/items/${fileId}`, rev: '"v2"', response: () => Response.json({ ...file, name: 'final.txt', eTag: '"v3"', parentReference: { id: folderId } }) },
    { name: 'delete_item', args: { item_id: fileId, etag: '"v2"' }, method: 'DELETE', path: `/me/drive/items/${fileId}`, rev: '"v2"', response: () => new Response(null, { status: 412 }), failed: true },
    { name: 'delete_item', args: { item_id: fileId, etag: '"v3"' }, method: 'DELETE', path: `/me/drive/items/${fileId}`, rev: '"v3"', response: () => new Response(null, { status: 204 }) },
  ];
  for (const operation of operations) {
    const seen = [];
    const { result } = await rpc('tools/call', { name: operation.name, arguments: operation.args }, async (path, init) => {
      seen.push({ path, method: init.method, version: new Headers(init.headers).get('If-Match') });
      return operation.response();
    });
    assert.deepEqual(seen.map(x => x.path.split('?')[0]), [operation.path]);
    assert.equal(seen[0].method, operation.method);
    if (operation.rev) assert.equal(seen[0].version, operation.rev);
    assert.equal(result.isError === true, operation.failed === true);
    if (operation.failed) assert.match(JSON.stringify(result.content), /412|version|changed/i);
  }
});


test('an oversized tool request is rejected before file access', async () => {
  const input = request('tools/call', { name: 'create_file', arguments: {} });
  input.headers.set('Content-Length', String(100 * 1024 * 1024));
  let calls = 0;
  const response = await handleMcp(input, {}, user, async () => { calls++; });
  assert.equal(response.status, 413);
  assert.equal(calls, 0);
  assert.match(await response.text(), /upload|attachment/i);
});
