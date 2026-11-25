import { getToken } from './ms-auth.mjs';
import { followResourceRedirect } from './transfer.mjs';

const graphOrigin = 'https://graph.microsoft.com';

export function graphClient(env, userId, followContent = false) {
  return async (path, init = {}) => {
    if (typeof path !== 'string' || !/^\/me\/drive(?:[/?]|$)/.test(path)) {
      throw new Error('Unsupported OneDrive request.');
    }
    const target = new URL('/v1.0' + path, graphOrigin);
    if (target.origin !== graphOrigin || !/^\/v1\.0\/me\/drive(?:\/|$)/.test(target.pathname)) {
      throw new Error('Unsupported OneDrive path.');
    }
    const access = await getToken(env, userId);
    const headers = new Headers(init.headers);
    headers.set('Authorization', 'Bearer ' + access);
    const method = (init.method || 'GET').toUpperCase();
    let response;
    try {
      response = await fetch(target, { ...init, method, headers, redirect: 'manual' });
    } catch {
      throw new Error('The request to OneDrive was interrupted. Check the item before retrying a change.');
    }
    if (followContent && method === 'GET' && target.pathname.endsWith('/content') && response.status >= 300 && response.status < 400) {
      return followResourceRedirect(response, init.headers || {});
    }
    return response;
  };
}
