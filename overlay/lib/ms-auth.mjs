const authority = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const scope = 'https://graph.microsoft.com/Files.ReadWrite offline_access';
const deviceGrant = 'urn:ietf:params:oauth:grant-type:device_code';
const now = () => Math.floor(Date.now() / 1000);
const clientPattern = /^[a-f\d]{8}(?:-[a-f\d]{4}){3}-[a-f\d]{12}$/i;
const encoder = new TextEncoder();
const pendingMessage = 'Start Microsoft sign-in from the private OneDrive setup page.';
const reconnectMessage = 'Connect your personal Microsoft account on the OneDrive setup page.';

export class AuthError extends Error {
  constructor(status, message) {
    super(message);
    this.status = status;
  }
}

function reply(data, status = 200) {
  return Response.json(data, {status, headers: {'Cache-Control': 'no-store', 'Referrer-Policy': 'no-referrer', 'X-Content-Type-Options': 'nosniff'}});
}

function useDatabase(env) {
  return env.DB.withSession ? env.DB.withSession('first-primary') : env.DB;
}

function toBase64(array) {
  return btoa(Array.from(array, x => String.fromCharCode(x)).join(''));
}

function fromBase64(value) {
  return Uint8Array.from(atob(value.replaceAll('-', '+').replaceAll('_', '/')), x => x.charCodeAt(0));
}

async function getKey(env) {
  try {
    const key = fromBase64(env.DRIVE_ENCRYPTION_KEY);
    if (key.length !== 32) throw new Error();
    return await crypto.subtle.importKey('raw', key, 'AES-GCM', false, ['encrypt','decrypt']);
  } catch {
    throw new AuthError(503, 'The OneDrive setup service requires its encryption key.');
  }
}

function aad(id, kind, clientId) {
  return encoder.encode(JSON.stringify([id, kind, clientId]));
}

async function encrypt(key, context, value) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const bytes = await crypto.subtle.encrypt({name: 'AES-GCM', iv, additionalData: context}, key, encoder.encode(JSON.stringify(value)));
  return toBase64(iv) + '.' + toBase64(new Uint8Array(bytes));
}

async function decrypt(env, row) {
  const key = await getKey(env);
  try {
    const [iv, cipher] = row.sealed.split('.');
    const plain = await crypto.subtle.decrypt({name: 'AES-GCM', iv: fromBase64(iv),
      additionalData: aad(row.user_id, row.kind, row.client_id)}, key, fromBase64(cipher));
    return JSON.parse(new TextDecoder().decode(plain));
  } catch {
    throw new AuthError(503, 'The saved connection could not be opened. Disconnect and connect again.');
  }
}

async function microsoft(endpoint, fields) {
  let response;
  try {
    response = await fetch(authority + '/' + endpoint, {
      method: 'POST',
      body: new URLSearchParams(fields),
      signal: AbortSignal.timeout(12000),
      // Workers supports manual redirect handling; OAuth credentials stay on this endpoint.
      redirect: 'manual',
    });
  } catch {
    throw new AuthError(503, 'The plugin could not reach Microsoft sign-in. Retry shortly.');
  }
  if (response.status >= 300 && response.status < 400) {
    await response.body?.cancel().catch(() => {});
    throw new AuthError(502, 'Microsoft redirected the sign-in request. The plugin stopped before sending credentials elsewhere.');
  }
  try {
    const body = await response.json();
    if (!body || typeof body !== 'object' || Array.isArray(body)) throw new TypeError();
    return {ok: response.ok, body, status: response.status};
  } catch {
    throw new AuthError(502, 'Microsoft sign-in returned an unreadable response. Retry shortly.');
  }
}

function verifiedAddress(value) {
  try {
    const url = new URL(value);
    const hosts = ['microsoft.com','www.microsoft.com','login.microsoftonline.com','login.live.com'];
    if (url.protocol === 'https:' && hosts.includes(url.hostname) && !url.port && !url.username && !url.password) return url.href;
  } catch { /* Use the Microsoft device sign-in address. */ }
  return 'https://www.microsoft.com/devicelogin';
}

async function rowFor(db, id, kind) {
  return db.prepare('SELECT * FROM onedrive_auth WHERE user_id = ? AND kind = ?').bind(id, kind).first();
}

async function removeRow(db, row) {
  return db.prepare('DELETE FROM onedrive_auth WHERE user_id = ? AND kind = ? AND revision = ?')
    .bind(row.user_id, row.kind, row.revision).run();
}

async function pendingDetails(env, row) {
  if (!row || !row.sealed || row.expires_at <= now()) return null;
  const data = await decrypt(env, row);
  return {userCode:data.userCode, verificationUri:data.verificationUri, expiresIn:Math.max(0,row.expires_at-now()),
    interval:Math.max(row.interval_s, row.next_at-now(), 1)};
}

async function status(env, db, id) {
  const rows = (await db.prepare('SELECT * FROM onedrive_auth WHERE user_id = ?').bind(id).all()).results;
  const active = rows.find(x => x.kind === 'active');
  const pending = rows.find(x => x.kind === 'pending');
  const details = await pendingDetails(env, pending);
  return {connected:!!active, pending:!!details, clientId:(details ? pending.client_id : active?.client_id) ?? null, ...details};
}

function newDeviceError(status) {
  if (status === 429 || status >= 500) return new AuthError(503, 'Microsoft sign-in is temporarily unavailable. Retry shortly.');
  return new AuthError(400, 'Microsoft could not start sign-in. Check the application (client) ID, personal account support, and Allow public client flows under Authentication.');
}

async function start(request, env, db, id) {
  let data;
  try { data = await request.json(); } catch { throw new AuthError(400, 'Enter a Microsoft application (client) ID.'); }
  if (!clientPattern.test(String(data?.clientId ?? ''))) throw new AuthError(400, 'Enter the 36-character Microsoft application (client) ID.');
  const clientId = data.clientId.toLowerCase();
  const key = await getKey(env);
  const revision = crypto.randomUUID();
  const row = {user_id:id, kind:'pending', revision};
  await db.prepare(`INSERT INTO onedrive_auth (user_id,kind,revision,client_id,sealed,expires_at,interval_s,next_at)
    VALUES (?,'pending',?,?,'',?,5,0) ON CONFLICT(user_id,kind) DO UPDATE SET
    revision=excluded.revision,client_id=excluded.client_id,sealed='',expires_at=excluded.expires_at,interval_s=5,next_at=0`)
    .bind(id, revision, clientId, now()+30).run();
  try {
    const result = await microsoft('devicecode', {client_id:clientId, scope});
    const {device_code, user_code, verification_uri, expires_in, interval} = result.body;
    if (!result.ok) throw newDeviceError(result.status);
    if (typeof device_code !== 'string' || !device_code || typeof user_code !== 'string' || !user_code || !Number.isSafeInteger(expires_in) || expires_in<=0)
      throw new AuthError(502,'Microsoft returned an incomplete sign-in response.');
    const intervalS = Math.max(5, Number.isSafeInteger(interval) ? interval : 5);
    const expiresIn = Math.min(expires_in, 1800);
    const verificationUri = verifiedAddress(verification_uri);
    const sealed = await encrypt(key, aad(id,'pending',clientId), {deviceCode:device_code, userCode:user_code, verificationUri});
    const write = await db.prepare(`UPDATE onedrive_auth SET sealed=?,expires_at=?,interval_s=?,next_at=?
       WHERE user_id=? AND kind='pending' AND revision=?`).bind(sealed,now()+expiresIn,intervalS,now()+intervalS,id,revision).run();
    if (!write.meta.changes) throw new AuthError(409,'A newer sign-in or disconnect replaced this request.');
    return {status:'pending',userCode:user_code, verificationUri, expiresIn, interval:intervalS};
  } catch (e) {
    await removeRow(db,row);
    throw e;
  }
}

function validatedTokens(body, priorRefresh) {
  const refresh = body.refresh_token || priorRefresh;
  if (String(body.token_type).toLowerCase() !== 'bearer' || typeof body.access_token !== 'string' ||
    !body.access_token || typeof refresh !== 'string' || !refresh || !Number.isSafeInteger(body.expires_in) || body.expires_in <= 0)
    throw new AuthError(502, 'Microsoft returned an incomplete token response. Reconnect your account.');
  if (typeof body.scope === 'string' && !body.scope.toLowerCase().split(/\s+/).some(x=> /^(https:\/\/graph\.microsoft\.com\/)?files\.readwrite(\.all)?$/.test(x)))
    throw new AuthError(403, 'Microsoft did not grant the requested OneDrive file permission. Reconnect and review the consent prompt.');
  return {access:body.access_token,refresh};
}

async function poll(env, db, id) {
  const row = await rowFor(db,id,'pending');
  if (!row) {
    if (await rowFor(db,id,'active')) return {status:'connected'};
    throw new AuthError(409,pendingMessage);
  }
  if (row.expires_at <= now()) {
    await removeRow(db,row);
    throw new AuthError(410,'The Microsoft sign-in code expired. Start sign-in again.');
  }
  if (!row.sealed || row.next_at > now()) return {status:'pending',interval:Math.max(1,row.next_at-now(),row.interval_s)};
  const key = await getKey(env);
  const session = await decrypt(env,row);
  // Reserve this poll across isolates, including the bounded upstream request time.
  const reserved = await db.prepare(`UPDATE onedrive_auth SET next_at=? WHERE user_id=? AND kind='pending' AND revision=? AND next_at<=?`)
    .bind(now()+Math.max(20,row.interval_s),id,row.revision,now()).run();
  if (!reserved.meta.changes) return {status:'pending', interval:row.interval_s};
  const result = await microsoft('token',{grant_type:deviceGrant,client_id:row.client_id,device_code:session.deviceCode});
  if (!result.ok) {
    const error = result.body.error;
    if (['authorization_pending','slow_down'].includes(error) || result.status===429 || result.status>=500) {
      const interval = row.interval_s + (error==='slow_down' || result.status===429 ? 5 : 0);
      await db.prepare(`UPDATE onedrive_auth SET next_at=?,interval_s=? WHERE user_id=? AND kind='pending' AND revision=?`)
        .bind(now()+interval,interval,id,row.revision).run();
      return {status:'pending', interval};
    }
    await removeRow(db,row);
    if (error==='authorization_declined' || error==='access_denied') throw new AuthError(403,'Microsoft sign-in was declined. Your existing connection is unchanged.');
    if (error==='expired_token' || error==='code_expired') throw new AuthError(410,'The Microsoft sign-in code expired. Start sign-in again.');
    throw new AuthError(400,'Microsoft could not complete sign-in. Verify that Allow public client flows is enabled, then start sign-in again.');
  }
  const token = validatedTokens(result.body);
  const sealed = await encrypt(key, aad(id,'active',row.client_id), token);
  const writes = await db.batch([
    db.prepare(`INSERT INTO onedrive_auth (user_id,kind,revision,client_id,sealed,expires_at,interval_s,next_at)
      SELECT user_id,'active',?,client_id,?,?,0,0 FROM onedrive_auth WHERE user_id=? AND kind='pending' AND revision=? AND expires_at>?
      ON CONFLICT(user_id,kind) DO UPDATE SET revision=excluded.revision, client_id=excluded.client_id,
      sealed=excluded.sealed, expires_at=excluded.expires_at, interval_s=0, next_at=0`)
      .bind(crypto.randomUUID(),sealed,now()+result.body.expires_in,id,row.revision,now()),
    db.prepare(`DELETE FROM drive_transfers WHERE owner=? AND EXISTS
      (SELECT 1 FROM onedrive_auth WHERE user_id=? AND kind='pending' AND revision=? AND expires_at>?)`)
      .bind(id,id,row.revision,now()),
    db.prepare(`DELETE FROM onedrive_auth WHERE user_id=? AND kind='pending' AND revision=?`).bind(id,row.revision)
  ]);
  if (!writes[0].meta.changes) throw new AuthError(409,'This sign-in is no longer current. Open the setup page again.');
  return {status:'connected'};
}

export async function handleAuth(request, env, userId) {
  try {
    if (!userId) throw new AuthError(401,'Sign in to this private Site.');
    const url = new URL(request.url);
    const action = url.pathname.slice('/api/connect/'.length);
    if (!['status','start','poll','disconnect'].includes(action)) throw new AuthError(404,'Unknown connection action.');
    if (request.method !== (action === 'status' ? 'GET' : 'POST')) throw new AuthError(405,'Unsupported request method.');
    if (request.method === 'POST' && request.headers.get('Origin') !== url.origin)
      throw new AuthError(403,'Open the private OneDrive setup page to change the connection.');
    const db = useDatabase(env);
    if (action === 'status') return reply(await status(env,db,userId));
    if (action === 'start') return reply(await start(request,env,db,userId));
    if (action === 'poll') return reply(await poll(env,db,userId));
    await db.batch([
      db.prepare('DELETE FROM onedrive_auth WHERE user_id = ?').bind(userId),
      db.prepare('DELETE FROM drive_transfers WHERE owner = ?').bind(userId)
    ]);
    return reply({status:'disconnected'});
  } catch (e) {
    return e instanceof AuthError ? reply({error:e.message},e.status) : reply({error:'OneDrive connection storage is temporarily unavailable.'},503);
  }
}

async function activeFor(db,id) {
  const row = await rowFor(db,id,'active');
  if (!row) throw new AuthError(401,reconnectMessage);
  return row;
}

export async function getToken(env, userId) {
  if (!userId) throw new AuthError(401, reconnectMessage);
  const db = useDatabase(env);
  // A refresh lease avoids duplicate calls; compare-and-swap also prevents a
  // late Microsoft response from restoring a disconnected or replaced account.
  for (let attempt=0; attempt<5; attempt++) {
    const row = await activeFor(db,userId);
    const token = await decrypt(env,row);
    if (row.expires_at > now()+60) return token.access;
    if (row.next_at > now()) {
      await new Promise(resolve=>setTimeout(resolve,300));
      continue;
    }
    const lease = crypto.randomUUID();
    const claim = await db.prepare(`UPDATE onedrive_auth SET next_at=?,revision=?
      WHERE user_id=? AND kind='active' AND revision=? AND next_at<=?`)
      .bind(now()+25,lease,userId,row.revision,now()).run();
    if (!claim.meta.changes) continue;
    try {
      const result = await microsoft('token',{grant_type:'refresh_token',client_id:row.client_id,refresh_token:token.refresh,scope});
      if (!result.ok) {
        if (['invalid_grant','interaction_required','consent_required','invalid_client','unauthorized_client'].includes(result.body.error)) {
          await removeRow(db,{...row, revision:lease});
          throw new AuthError(401,reconnectMessage);
        }
        throw new AuthError(503,'Microsoft could not renew the connection. Retry shortly.');
      }
      const updated = validatedTokens(result.body,token.refresh);
      const sealed = await encrypt(await getKey(env),aad(userId,'active',row.client_id),updated);
      const write = await db.prepare(`UPDATE onedrive_auth SET sealed=?,expires_at=?,next_at=0,revision=?
        WHERE user_id=? AND kind='active' AND revision=?`).bind(sealed,now()+result.body.expires_in,crypto.randomUUID(),userId,lease).run();
      if (!write.meta.changes) throw new AuthError(409,'The Microsoft connection changed. Retry the operation.');
      return updated.access;
    } catch (e) {
      await db.prepare(`UPDATE onedrive_auth SET next_at=0 WHERE user_id=? AND kind='active' AND revision=?`).bind(userId,lease).run();
      throw e;
    }
  }
  throw new AuthError(503,'Microsoft sign-in is being renewed. Retry shortly.');
}
