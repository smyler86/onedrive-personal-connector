import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { Miniflare } from 'miniflare';

const root = new URL('../', import.meta.url);
const site = 'https://onedrive.example';
const clientId = '12345678-1234-1234-1234-123456789abc';
const endpoint = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode';

async function fixture(t, providerResponse) {
  const requests = [];
  const runtime = new Miniflare({
    compatibilityDate: '2026-05-15',
    compatibilityFlags: ['nodejs_compat'],
    modulesRoot: fileURLToPath(root),
    modules: [
      {
        type: 'ESModule',
        path: fileURLToPath(new URL('sign-in-test-worker.mjs', root)),
        contents: `import { handleAuth } from './lib/ms-auth.mjs';
          export default { fetch(request, env) { return handleAuth(request, env, 'test-user'); } };`,
      },
      {
        type: 'ESModule',
        path: fileURLToPath(new URL('lib/ms-auth.mjs', root)),
        contents: await readFile(new URL('lib/ms-auth.mjs', root), 'utf8'),
      },
    ],
    d1Databases: ['DB'],
    bindings: { DRIVE_ENCRYPTION_KEY: Buffer.alloc(32, 7).toString('base64') },
    outboundService: async request => {
      requests.push({ url: request.url, method: request.method, type: request.headers.get('content-type'), body: await request.text() });
      return providerResponse();
    },
  });
  t.after(() => runtime.dispose());
  const db = await runtime.getD1Database('DB');
  const migrations = new URL('drizzle/', root);
  for (const file of (await readdir(migrations)).filter(name => name.endsWith('.sql')).sort()) {
    const sql = await readFile(new URL(file, migrations), 'utf8');
    for (const statement of sql.split('--> statement-breakpoint').map(s => s.trim()).filter(Boolean)) {
      await db.prepare(statement).run();
    }
  }
  const start = () => runtime.dispatchFetch(site + '/api/connect/start', {
    method: 'POST',
    headers: { Origin: site, 'Content-Type': 'application/json' },
    body: JSON.stringify({clientId}),
  });
  return { db, requests, start };
}

test('the production Worker sends the form to Microsoft and returns a sign-in code', async t => {
  const providerReply = {
    device_code: 'internal-device-credential',
    user_code: 'SAMPLE-42',
    verification_uri: 'https://www.microsoft.com/devicelogin',
    expires_in: 900,
    interval: 5,
  };
  const { db, requests, start } = await fixture(t, () => Response.json(providerReply));
  const response = await start();
  assert.equal(response.status, 200);
  const result = await response.json();
  assert.equal(result.status, 'pending');
  assert.equal(result.userCode, providerReply.user_code);
  assert.equal(result.verificationUri, providerReply.verification_uri);
  assert.ok(!JSON.stringify(result).includes(providerReply.device_code));
  assert.equal(requests.length, 1);
  assert.equal(requests[0].url, endpoint);
  assert.equal(requests[0].method, 'POST');
  assert.match(requests[0].type, /^application\/x-www-form-urlencoded\b/);
  const form = new URLSearchParams(requests[0].body);
  assert.equal(form.get('client_id'), clientId);
  assert.equal(form.get('scope'), 'https://graph.microsoft.com/Files.ReadWrite offline_access');
  const pending = await db.prepare("SELECT sealed FROM onedrive_auth WHERE user_id = ? AND kind = 'pending'").bind('test-user').first();
  assert.ok(pending?.sealed);
  assert.ok(!pending.sealed.includes(providerReply.device_code));
});

test('the production Worker stops an OAuth redirect and clears incomplete sign-in', async t => {
  const location = 'https://elsewhere.example/oauth?private=fixture';
  const { db, requests, start } = await fixture(t, () => Response.redirect(location, 307));
  const response = await start();
  assert.equal(response.status, 502);
  const result = await response.json();
  assert.match(result.error, /redirect/i);
  assert.ok(!result.error.includes(location));
  assert.deepEqual(requests.map(request => request.url), [endpoint]);
  assert.equal(await db.prepare('SELECT COUNT(*) AS n FROM onedrive_auth').first('n'), 0);
});
