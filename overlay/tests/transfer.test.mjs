import test from 'node:test';
import assert from 'node:assert/strict';
import { DatabaseSync } from 'node:sqlite';
import { readdirSync, readFileSync } from 'node:fs';
import { callTransferTool, followResourceRedirect, handleTransfer } from '../lib/transfer.mjs';

const origin = 'https://personal.test';
const upload = 'https://my.microsoftpersonalcontent.com/personal/test/upload?ticket=only-for-server';
const endpoint = '/api/files/upload/';
const three = [3,4,5];
const rev = '"source-v1"';
const getExpiry = () => new Date(Date.now()+300000).toISOString();
const response = (obj, code=200) => Response.json(obj,{status:code});
function setup(t) {
  const sqlite = new DatabaseSync(':memory:');
  t.after(()=>sqlite.close());
  const migrations = new URL('../drizzle/',import.meta.url);
  for (const path of readdirSync(migrations).filter(p=>p.endsWith('.sql')).sort()) {
    sqlite.exec(readFileSync(new URL(path,migrations),'utf8'));
  }
  const DB = {prepare(sql) {return {bind(...args) {
    const stmt = sqlite.prepare(sql);
    return {async run() {return stmt.run(...args);}, async first() {return stmt.get(...args);}};
  }};}};
  const env = {DB,DRIVE_ENCRYPTION_KEY:Buffer.alloc(32,0x32).toString('base64')};
  const state = {drive:'personal-drive-A', rev, saved:[], calls:[], outgoing:[], batches:[], owner:'u', offset:0, structured:true};
  const active = () => ({expirationDateTime:getExpiry(),nextExpectedRanges:[String(state.offset)+'-']});
  state.graph = async (path, init={}) => {
    state.calls.push({path,init});
    if (path === '/me/drive?$select=id') return response({id:state.drive});
    if (path.includes('/createUploadSession')) {
      state.start = JSON.parse(init.body);
      return response({uploadUrl:upload,...active()});
    }
    if (path.includes('?$select=id,name,eTag,file')) return response({id:'i2',name:'existing.json',eTag:state.rev,file:{}});
    if (path.includes('?$select=id,eTag')) return response({id:'i2',eTag:state.rev});
    if (path.includes('?$select=id,name,file')) return response({id:'i2',name:'example.html',file:{}});
    if (init.method === 'PUT' && path.includes('/content')) return response({id:'i2',name:'empty',size:0,eTag:"e"},201);
    return new Response(null,{status:302, headers:{location:'https://mock.files.1drv.com/f?secret=hidden'}});
  };
  state.wire = async (target, opts={}) => {
    state.outgoing.push({target:target.toString(),opts});
    assert.equal(new Headers(opts.headers).has('authorization'),false);
    assert.equal(opts.redirect,'manual');
    if (target.hostname.endsWith('files.1drv.com')) {
      return new Response('yes',{status:206,headers:{'content-type':'text/html','content-range':'bytes 1-3/9','content-length':'3'}});
    }
    assert.equal(target.toString(),upload);
    if (state.fromNetwork) return state.fromNetwork(opts);
    if (opts.method === 'DELETE') return new Response(null,{status:200});
    if (opts.method === 'GET') return response(active());
    const match = /bytes (\d+)-(\d+)\/(\d+)/.exec(new Headers(opts.headers).get('content-range'));
    if (+match[1] !== state.offset) return response({},416);
    state.offset = +match[2]+1;
    state.batches.push(opts.body.length);
    if (state.offset < +match[3]) return response(active(),202);
    state.saved.push(1);
    return response({id:'i2',name:'returned-file',size:state.offset,eTag:'v2','@microsoft.graph.downloadUrl':'https://excluded?secret'},201);
  };
  state.send = async (method, path, body, headers={}, user=state.owner) => {
    const opts = {method,headers:{origin,...headers}};
    if (body !== undefined) {
      opts.body = typeof body === 'string' || body instanceof Uint8Array ? body : JSON.stringify(body);
      if (!(body instanceof Uint8Array) && !opts.headers['content-type']) opts.headers['content-type']='application/json';
    }
    return handleTransfer(new Request(origin+path,opts),env,user,state.graph,state.wire);
  };
  state.begin = async (args) => {
    const res = await state.send('POST',endpoint+'start',args);
    assert.equal(res.status,200);
    return res.json();
  };
  state.sid = (suffix,id) => endpoint+suffix+'?sessionId='+id;
  state.chunk = (id, range, input) => state.send('PUT',state.sid('chunk',id),input,{'content-range':range});
  state.sqlite=sqlite;
  return state;
}
test('chunked upload verifies item output and allows a post-commit status query and re-delivery',async t=>{
  const s=setup(t);
  const chunk = 320*1024; const total=chunk+3;
  const {sessionId,chunkSize} = await s.begin({name:'Photo and video.mov',size:total,parentId:'p1'});
  assert.ok(chunkSize >= chunk);
  assert.match(s.calls[1].path,/p1:\/Photo%20and%20video.mov:\/createUploadSession$/);
  assert.equal(s.start.item['@microsoft.graph.conflictBehavior'],'fail');
  const vault=s.sqlite.prepare('SELECT session_json FROM drive_transfers').get().session_json;
  assert.doesNotMatch(vault,/personal|ticket/);
  const first = await s.chunk(sessionId,'bytes 0-327679/'+total,new Uint8Array(chunk));
  assert.deepEqual(await first.json(),{complete:false,offset:chunk});
  const status = await s.send('GET',s.sid('status',sessionId));
  assert.equal((await status.json()).offset,chunk);
  const completed = await s.chunk(sessionId,'bytes 327680-327682/'+total,Uint8Array.from(three));
  assert.equal(completed.status,200);
  const ack=await completed.json();
  assert.equal(ack.complete,true);
  assert.equal(ack.item.id,'i2');
  assert.doesNotMatch(JSON.stringify(ack),/secret|ticket|downloadUrl/);
  const completeCheck=await s.send('GET',s.sid('status',sessionId));
  assert.deepEqual(await completeCheck.json(),ack);
  const repeated=await s.chunk(sessionId,'bytes 327680-327682/'+total,Uint8Array.from(three));
  assert.deepEqual(await repeated.json(),ack);
  assert.equal(s.saved.length,1);
});
test('user/connection isolation and offset/target validation',async t=>{
  const s=setup(t);
  const {sessionId}=await s.begin({name:'own.bin',size:3});
  const noUser=await s.send('GET',s.sid('status',sessionId),undefined,{},'');
  assert.equal(noUser.status,401);
  assert.equal((await s.send('GET',s.sid('status',sessionId),undefined,{},'other')).status,404);
  s.drive='different';
  assert.equal((await s.send('GET',s.sid('status',sessionId))).status,403);
  s.drive='personal-drive-A';
  assert.equal((await s.chunk(sessionId,'bytes 1-2/3',Uint8Array.from([3,3]))).status,400);
  assert.equal((await s.chunk(sessionId,'bytes 0-2/3',Uint8Array.from([3]))).status,400);
  assert.equal((await s.send('POST',endpoint+'start',{name:'bad',size:1},{origin:'https://wrong'})).status,403);
  const manual = await s.send('POST',endpoint+'start',{name:'yourfile',size:1,url:'https://example.com'});
  assert.equal(manual.status,400);
  assert.equal(s.outgoing.length,0);
  assert.equal((await s.send('DELETE',s.sid('session',sessionId))).status,200);
  assert.equal(s.outgoing.at(-1).opts.method,'DELETE');
  assert.equal((await s.send('GET',s.sid('status',sessionId))).status,404);
});
test('replacing a file calls the versioned target and rejects changes detected during the transfer',async t=>{
  const s=setup(t);
  const {sessionId} = await s.begin({itemId:'i2',eTag:rev,name:'different.mp4',size:3});
  const authSession = s.calls.find(x=>x.path.endsWith('/createUploadSession'));
  assert.equal(authSession.init.headers['If-Match'],rev);
  assert.equal(s.start.item.name,'existing.json');
  s.rev='"new-version"';
  assert.equal((await s.chunk(sessionId,'bytes 0-2/3',Uint8Array.from(three))).status,412);
  assert.equal(s.outgoing.length,0);
  assert.equal((await s.send('POST',endpoint+'start',{itemId:'i2',eTag:rev,size:0})).status,412);
});
test('an empty file is created without losing conflict protection',async t=>{
  const s=setup(t);
  const ack=await s.begin({name:'empty',size:0});
  assert.equal(ack.complete,true);
  assert.equal(s.calls.at(-1).path,'/me/drive/root:/empty:/content?@microsoft.graph.conflictBehavior=fail');
  assert.equal(s.calls.at(-1).init.body.byteLength,0);
});
test('a missing final acknowledgement is returned as uncertain, without creating a new session',async t=>{
  const s=setup(t);
  const {sessionId} = await s.begin({name:'test.odt',size:3});
  s.fromNetwork = () => response({},404);
  const stale=await s.send('GET',s.sid('status',sessionId));
  assert.equal(stale.status,409);
  assert.match((await stale.json()).error,/Verify whether/);
  assert.equal(s.calls.filter(c=>c.path.includes('createUploadSession')).length,1);
});
test('download has a local agent-controlled endpoint, authenticated resource, range semantics and no active-document render',async t=>{
  const s=setup(t);
  const [link] = await callTransferTool('get_download_link',{itemId:'i2'},origin);
  assert.equal(JSON.parse(link.text).url,origin+'/api/files/download?id=i2');
  const clip = await s.send('GET','/api/files/download?id=i2',undefined,{range:'bytes=1-3'});
  assert.equal(clip.status,206);
  assert.equal(s.outgoing.length,1);
  assert.equal(clip.headers.get('content-type'),'application/octet-stream');
  assert.match(clip.headers.get('content-disposition'),/^attachment/);
  assert.equal(clip.headers.get('x-content-type-options'),'nosniff');
  assert.equal(clip.headers.get('content-range'),'bytes 1-3/9');
  assert.equal(new Headers(s.outgoing[0].opts.headers).get('range'),'bytes=1-3');
  assert.equal(await clip.text(),'yes');
});
test('the host helper strips authorization and does not attempt unapproved hops',async ()=>{
  const upstream=new Response(null,{status:302,headers:{location:'https://one.files.1drv.com/content'}});
  let calls = 0;
  await assert.rejects(followResourceRedirect(upstream,{Authorization:'Bearer internal',Cookie:'bad',Range:'bytes=20-'},async (target,init)=>{
    calls++;
    assert.deepEqual(Object.fromEntries(new Headers(init.headers)),{range:'bytes=20-'});
    return new Response(null,{status:302,headers:{location:'https://one.files.1drv.com.example/handoff'}});
  }),/unsupported transfer endpoint/);
  assert.equal(calls,1);
});
test('external session addresses reject non-Microsoft hosts and stale parent arguments never reach the wire',async t=>{
  const s=setup(t);
  s.graph=async (path)=> path.endsWith('createUploadSession') ? response({uploadUrl:'https://example.com/collect',expirationDateTime:getExpiry()}) : response({id:'root'});
  assert.equal((await s.send('POST',endpoint+'start',{name:'A',size:2})).status,502);
  assert.equal(s.outgoing.length,0);
  await assert.rejects(callTransferTool('open_upload_page',{itemId:'one',eTag:'*'},origin),/eTag/);
  await assert.rejects(callTransferTool('open_upload_page',{itemId:'one',eTag:rev,parentId:'root'},origin),/replacement/);
});
