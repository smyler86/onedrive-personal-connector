import test from 'node:test';
import assert from 'node:assert/strict';
import {DatabaseSync} from 'node:sqlite';
import {readdirSync, readFileSync} from 'node:fs';
import {handleAuth, getToken} from '../lib/ms-auth.mjs';

const origin = 'https://onedrive.example';
const clientId = '12345678-1234-1234-1234-123456789abc';
const device = {device_code:'internal-device-code', user_code:'SHORT123', verification_uri:'https://www.microsoft.com/devicelogin', expires_in:900,interval:5};
const tokens = {access_token:'test-access-value',refresh_token:'test-refresh-value',expires_in:3600,scope:'Files.ReadWrite',token_type:'Bearer'};
const json = (data,status=200) => Response.json(data,{status});

function fixture(t) {
  const sqlite = new DatabaseSync(':memory:');
  t.after(()=>sqlite.close());
  const migrations = new URL('../drizzle/',import.meta.url);
  for (const path of readdirSync(migrations).filter(x=>x.endsWith('.sql')).sort())
    sqlite.exec(readFileSync(new URL(path,migrations),'utf8'));
  const DB = {
    prepare(query) {
      return {bind(...params) {
        const statement = sqlite.prepare(query);
        return {
          async first() { return statement.get(...params); },
          async all() { return {results:statement.all(...params)}; },
          async run() { return {meta:{changes:Number(statement.run(...params).changes)}}; }
        };
      }};
    },
    async batch(statements) {
      sqlite.exec('BEGIN');
      try {
        const results = [];
        for (const statement of statements) results.push(await statement.run());
        sqlite.exec('COMMIT');
        return results;
      } catch (error) {
        sqlite.exec('ROLLBACK');
        throw error;
      }
    }
  };
  const env = {DB,DRIVE_ENCRYPTION_KEY:Buffer.alloc(32,42).toString('base64')};
  function request(action,body,user='a',pageOrigin=origin) {
    const init = action === 'status' ? {} : {method:'POST',headers:{Origin:pageOrigin,'Content-Type':'application/json'},body:JSON.stringify(body ?? {})};
    return handleAuth(new Request(origin+'/api/connect/'+action,init),env,user);
  }
  const pollNow = (user='a')=>sqlite.prepare("UPDATE onedrive_auth SET next_at=0 WHERE user_id=? AND kind='pending'").run(user);
  const activeExpires = (user='a')=>sqlite.prepare("UPDATE onedrive_auth SET expires_at=0 WHERE user_id=? AND kind='active'").run(user);
  const addTransfer = (id,owner)=>sqlite.prepare('INSERT INTO drive_transfers(id,owner,session_json,expires_at) VALUES(?,?,?,?)').run(id,owner,'test',500);
  const countTransfer = owner=>sqlite.prepare('SELECT COUNT(*) AS n FROM drive_transfers WHERE owner=?').get(owner).n;
  async function connect(user='a') {
    assert.equal((await request('start',{clientId},user)).status,200);
    pollNow(user);
    assert.deepEqual(await (await request('poll',{},user)).json(),{status:'connected'});
  }
  return {sqlite,env,request,pollNow,activeExpires,connect,addTransfer,countTransfer};
}

test('Microsoft sign-in isolates pending and active secrets, enforces origin and limits Microsoft scopes',async t=>{
  const f=fixture(t);
  let requests=0;
  t.mock.method(globalThis,'fetch',async (url,opts)=>{
    requests++;
    assert.match(url,/^https:\/\/login\.microsoftonline\.com\/consumers\/oauth2\/v2\.0\/(devicecode|token)$/);
    assert.equal(opts.body.get('client_id'),clientId);
    if (url.endsWith('/devicecode')) {
      assert.equal(opts.body.get('scope'),'https://graph.microsoft.com/Files.ReadWrite offline_access');
      return json(device);
    }
    assert.equal(opts.body.get('device_code'),device.device_code);
    return json(tokens);
  });
  assert.equal((await f.request('start',{clientId},'a','https://wrong.example')).status,403);
  assert.equal((await f.request('status',{},null)).status,401);
  assert.equal((await f.request('start',{clientId:'test'})).status,400);
  assert.equal(requests,0);
  const started=await (await f.request('start',{clientId})).json();
  assert.equal(started.userCode,device.user_code);
  assert.equal(started.verificationUri,device.verification_uri);
  assert.ok(!JSON.stringify(started).includes(device.device_code));
  assert.equal((await (await f.request('status')).json()).userCode,device.user_code);
  assert.equal((await (await f.request('status',{},'b')).json()).pending,false);
  assert.equal((await f.request('poll',{},'b')).status,409);
  assert.ok(!JSON.stringify(f.sqlite.prepare('SELECT * FROM onedrive_auth').get()).includes(device.user_code));
  assert.equal((await (await f.request('poll')).json()).status,'pending');
  assert.equal(requests,1);
  f.pollNow();
  assert.equal((await (await f.request('poll')).json()).status,'connected');
  assert.equal(requests,2);
  assert.equal(await getToken(f.env,'a'),tokens.access_token);
  await assert.rejects(getToken(f.env,'b'),{status:401});
  const stored=JSON.stringify(f.sqlite.prepare('SELECT * FROM onedrive_auth').get());
  assert.ok(!stored.includes(tokens.access_token));
  assert.ok(!stored.includes(tokens.refresh_token));
  assert.equal((await (await f.request('poll')).json()).status,'connected');
  assert.equal(requests,2);
});

test('slow polling, expiry, and denied replacements keep the previous authorized account',async t=>{
  const f=fixture(t);
  let response=()=>json(tokens);
  t.mock.method(globalThis,'fetch',async url=>url.endsWith('devicecode') ? json(device) : response());
  await f.connect();
  await f.request('start',{clientId});
  response = ()=>json({error:'slow_down'},400);
  f.pollNow();
  assert.deepEqual(await (await f.request('poll')).json(),{status:'pending',interval:10});
  f.pollNow();
  response = ()=>json({error:'authorization_declined',error_description:'sensitive upstream value'},400);
  const declined=await f.request('poll');
  assert.equal(declined.status,403);
  assert.doesNotMatch(await declined.text(),/sensitive/);
  assert.equal(await getToken(f.env,'a'),tokens.access_token);
  await f.request('start',{clientId});
  f.sqlite.exec("UPDATE onedrive_auth SET expires_at=0 WHERE kind='pending'");
  assert.equal((await f.request('poll')).status,410);
  assert.equal((await (await f.request('status')).json()).pending,false);
});

test('disconnect or account replacement cancels transfer handles only for the caller',async t=>{
  const f=fixture(t);
  let tokenResponse=()=>json(tokens);
  t.mock.method(globalThis,'fetch',async url=>url.endsWith('devicecode') ? json(device) : tokenResponse());
  await f.connect();
  f.addTransfer('a1','a');
  f.addTransfer('b1','b');
  await f.connect();
  assert.equal(f.countTransfer('a'),0);
  assert.equal(f.countTransfer('b'),1);
  f.addTransfer('a2','a');
  await f.request('start',{clientId});
  f.pollNow();
  let release;
  let started;
  const entered=new Promise(r=>{started=r;});
  tokenResponse=()=> new Promise(resolve=>{release=()=>resolve(json(tokens));started();});
  const poll = f.request('poll');
  await entered;
  assert.equal((await f.request('disconnect')).status,200);
  release();
  assert.equal((await poll).status,409);
  await assert.rejects(getToken(f.env,'a'),{status:401});
  assert.equal(f.countTransfer('a'),0);
  assert.equal(f.countTransfer('b'),1);
});

test('a concurrent refresh is coalesced, rotates secrets, and invalid_grant requires new sign-in',async t=>{
  const f=fixture(t);
  let refreshCount=0;
  let invalid=false;
  t.mock.method(globalThis,'fetch',async (url,opts)=>{
    if (url.endsWith('devicecode')) return json(device);
    if (opts.body.get('grant_type') !== 'refresh_token') return json(tokens);
    const refresh = opts.body.get('refresh_token');
    refreshCount++;
    if (invalid) {
      assert.equal(refresh,'rotated-refresh');
      return json({error:'invalid_grant'},400);
    }
    assert.equal(refresh,tokens.refresh_token);
    await new Promise(r=>setTimeout(r,20));
    return json({...tokens,access_token:'rotated-access',refresh_token:'rotated-refresh'});
  });
  await f.connect();
  f.activeExpires();
  assert.deepEqual(await Promise.all([getToken(f.env,'a'),getToken(f.env,'a')]),['rotated-access','rotated-access']);
  assert.equal(refreshCount,1);
  f.activeExpires();
  invalid=true;
  await assert.rejects(getToken(f.env,'a'),{status:401});
  assert.equal((await (await f.request('status')).json()).connected,false);
});

test('late refresh response cannot restore a disconnected account',async t=>{
  const f=fixture(t);
  let release;
  let entered;
  const started=new Promise(r=>{entered=r;});
  t.mock.method(globalThis,'fetch',async (url,opts)=>{
    if (url.endsWith('devicecode')) return json(device);
    if (opts.body.get('grant_type')!=='refresh_token') return json(tokens);
    return new Promise(r=>{entered();release=()=>r(json(tokens));});
  });
  await f.connect();
  f.activeExpires();
  const renew=getToken(f.env,'a');
  await started;
  await f.request('disconnect');
  release();
  await assert.rejects(renew,{status:409});
  await assert.rejects(getToken(f.env,'a'),{status:401});
});
