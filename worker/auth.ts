import type { Env } from './graph';

const SCOPES = 'https://graph.microsoft.com/User.Read https://graph.microsoft.com/Presence.Read https://graph.microsoft.com/Calendars.ReadBasic https://graph.microsoft.com/MailboxSettings.Read offline_access';
const TOKEN_KEY = 'microsoft-connection-v1';
const COOKIE = '__Host-office-oauth';
export class ReconnectRequired extends Error { constructor() { super('Microsoft connection requires sign-in'); } }
interface Tokens { accessToken: string; refreshToken: string; expiresAt: number; userId: string; target: string; clientId: string; tenantId: string }
interface Flow { state: string; verifier: string; expiresAt: number }
function encode(bytes: Uint8Array) { return btoa(String.fromCharCode(...bytes)).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,''); }
function decode(value: string) { return Uint8Array.from(atob(value.replace(/-/g,'+').replace(/_/g,'/')), c=>c.charCodeAt(0)); }
function random() { return encode(crypto.getRandomValues(new Uint8Array(32))); }
async function key(env: Env) {
  if (!env.TOKEN_ENCRYPTION_KEY) throw new Error('Missing token encryption key');
  const bytes = decode(env.TOKEN_ENCRYPTION_KEY);
  if (bytes.length !== 32) throw new Error('Token encryption key must be 32 bytes');
  return crypto.subtle.importKey('raw',bytes,{name:'AES-GCM'},false,['encrypt','decrypt']);
}
export async function encrypt(value: unknown, env: Env): Promise<string> {
  const iv=crypto.getRandomValues(new Uint8Array(12));
  const ciphertext=await crypto.subtle.encrypt({name:'AES-GCM',iv},await key(env),new TextEncoder().encode(JSON.stringify(value)));
  return `${encode(iv)}.${encode(new Uint8Array(ciphertext))}`;
}
export async function decrypt<T>(value: string, env: Env): Promise<T> {
  const [iv,encrypted]=value.split('.');
  const plaintext=await crypto.subtle.decrypt({name:'AES-GCM',iv:decode(iv)},await key(env),decode(encrypted));
  return JSON.parse(new TextDecoder().decode(plaintext)) as T;
}
function config(env:Env) {
  if(!env.MICROSOFT_TENANT_ID || !env.MICROSOFT_CLIENT_ID || !env.MICROSOFT_CLIENT_SECRET || !env.TARGET_USER_ID || !env.TOKEN_ENCRYPTION_KEY || !env.CONNECTIONS || !env.APP_ORIGIN) throw new Error('Microsoft sign-in is not configured');
  const origin=new URL(env.APP_ORIGIN);
  if(origin.protocol!=='https:' || origin.origin!==env.APP_ORIGIN) throw new Error('APP_ORIGIN must be an HTTPS origin');
  return `${env.APP_ORIGIN}/api/auth/callback`;
}
async function exchange(env:Env, fields:Record<string,string>) {
  const response=await fetch(`https://login.microsoftonline.com/${encodeURIComponent(env.MICROSOFT_TENANT_ID!)}/oauth2/v2.0/token`,{method:'POST',body:new URLSearchParams({client_id:env.MICROSOFT_CLIENT_ID!,client_secret:env.MICROSOFT_CLIENT_SECRET!,scope:SCOPES,...fields}),signal:AbortSignal.timeout(10000)});
  if(!response.ok) {
    const error=await response.json().catch(()=>({})) as {error?:string};
    if(fields.grant_type==='refresh_token' && ['invalid_grant','interaction_required'].includes(error.error || '')) throw new ReconnectRequired();
    throw new Error(`Microsoft sign-in failed (${response.status})`);
  }
  const data=await response.json() as {access_token:string;refresh_token?:string;expires_in:number};
  if(!data.access_token || !Number.isFinite(data.expires_in)) throw new Error('Invalid Microsoft token response');
  return data;
}
let refreshing: Promise<Tokens> | undefined;
export async function delegatedToken(env:Env):Promise<string> {
  config(env);
  const stored=await env.CONNECTIONS!.get(TOKEN_KEY);
  if(!stored) throw new ReconnectRequired();
  const tokens=await decrypt<Tokens>(stored,env);
  if(tokens.target!==env.TARGET_USER_ID!.toLowerCase() || tokens.clientId!==env.MICROSOFT_CLIENT_ID || tokens.tenantId!==env.MICROSOFT_TENANT_ID) throw new ReconnectRequired();
  if(tokens.expiresAt>Date.now()+60000) return tokens.accessToken;
  // Deduplicate refreshes within each Worker isolate. Microsoft permits old refresh
  // tokens to remain valid; this single-account KV implementation is not a distributed lock.
  refreshing ??= (async()=>{
    const data=await exchange(env,{grant_type:'refresh_token',refresh_token:tokens.refreshToken});
    const next={...tokens,accessToken:data.access_token,refreshToken:data.refresh_token || tokens.refreshToken,expiresAt:Date.now()+data.expires_in*1000};
    await env.CONNECTIONS!.put(TOKEN_KEY,await encrypt(next,env));return next;
  })();
  try { return (await refreshing).accessToken; } finally { refreshing=undefined; }
}
const security={'Cache-Control':'no-store','Referrer-Policy':'no-referrer','X-Content-Type-Options':'nosniff','Content-Security-Policy':"default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'"};
function page(message:string,status:number) {
  return new Response(`<!doctype html><html lang="en"><meta name="viewport" content="width=device-width"><title>Microsoft connection</title><body style="background:#101915;color:#f5f7f1;font:20px system-ui;padding:48px;max-width:680px"><h1>Microsoft connection</h1><p>${message}</p><p><a style="color:#b4f5a3" href="/connect">Back to connection settings</a></p></body></html>`,{status,headers:{...security,'Content-Type':'text/html; charset=utf-8','Set-Cookie':`${COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`}});
}
export async function authRoute(request:Request,env:Env):Promise<Response> {
  const url=new URL(request.url);
  if(url.pathname==='/api/auth/connection') {
    if(request.method!=='GET') return new Response('Method not allowed',{status:405,headers:security});
    let connected=false;
    try {const stored=await env.CONNECTIONS?.get(TOKEN_KEY);if(stored){const data=await decrypt<Tokens>(stored,env);connected=data.target===env.TARGET_USER_ID?.toLowerCase() && data.clientId===env.MICROSOFT_CLIENT_ID && data.tenantId===env.MICROSOFT_TENANT_ID;}}catch{ /* Report safe connection state only. */ }
    return Response.json({connected,configured:!!(env.MICROSOFT_CLIENT_SECRET && env.TOKEN_ENCRYPTION_KEY && env.TARGET_USER_ID && env.CONNECTIONS),demo:env.DEMO_MODE==='true'},{headers:security});
  }
  try {
    const redirectUri=config(env);
    if(url.origin!==env.APP_ORIGIN) return page('Open connection settings on the configured application hostname.',400);
    if(url.pathname==='/api/auth/start') {
      if(request.method!=='POST') return new Response('Method not allowed',{status:405,headers:{...security,Allow:'POST'}});
      if(request.headers.get('Origin')!==env.APP_ORIGIN) return page('Invalid connection request. Please try again from the connection page.',403);
      const flow:Flow={state:random(),verifier:random(),expiresAt:Date.now()+600000};
      const challenge=encode(new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode(flow.verifier))));
      const query=new URLSearchParams({client_id:env.MICROSOFT_CLIENT_ID!,response_type:'code',redirect_uri:redirectUri,response_mode:'query',scope:SCOPES,state:flow.state,code_challenge:challenge,code_challenge_method:'S256',prompt:'select_account'});
      return new Response(null,{status:303,headers:{...security,Location:`https://login.microsoftonline.com/${encodeURIComponent(env.MICROSOFT_TENANT_ID!)}/oauth2/v2.0/authorize?${query}`,'Set-Cookie':`${COOKIE}=${await encrypt(flow,env)}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=600`}});
    }
    if(url.pathname==='/api/auth/callback') {
      if(request.method!=='GET') return new Response('Method not allowed',{status:405,headers:security});
      const cookie=request.headers.get('Cookie')?.split(';').map(c=>c.trim()).find(c=>c.startsWith(`${COOKIE}=`))?.slice(COOKIE.length+1);
      if(!cookie) return page('Sign-in session expired. Please start again.',400);
      const flow=await decrypt<Flow>(cookie,env);
      if(flow.expiresAt<Date.now() || !flow.state || url.searchParams.get('state')!==flow.state) return page('Sign-in session is invalid or expired. Please start again.',400);
      const code=url.searchParams.get('code');
      if(!code || url.searchParams.has('error')) return page('Microsoft sign-in was cancelled or denied. Your existing connection has not changed.',400);
      const data=await exchange(env,{grant_type:'authorization_code',code,redirect_uri:redirectUri,code_verifier:flow.verifier});
      if(!data.refresh_token) throw new Error('Microsoft did not issue offline access');
      const response=await fetch('https://graph.microsoft.com/v1.0/me?$select=id,userPrincipalName,mail',{headers:{Authorization:`Bearer ${data.access_token}`},signal:AbortSignal.timeout(10000)});
      if(!response.ok) throw new Error(`Microsoft profile request failed (${response.status})`);
      const user=await response.json() as {id:string;userPrincipalName:string;mail?:string};
      const target=env.TARGET_USER_ID!.toLowerCase();
      if(![user.id,user.userPrincipalName,user.mail].some(value=>value?.toLowerCase()===target)) return page('This account is not the configured office user. Sign in with the correct Microsoft 365 account.',403);
      const tokens:Tokens={accessToken:data.access_token,refreshToken:data.refresh_token,expiresAt:Date.now()+data.expires_in*1000,userId:user.id,target,clientId:env.MICROSOFT_CLIENT_ID!,tenantId:env.MICROSOFT_TENANT_ID!};
      await env.CONNECTIONS!.put(TOKEN_KEY,await encrypt(tokens,env));
      return new Response(null,{status:303,headers:{...security,Location:'/connect?connected=1','Set-Cookie':`${COOKIE}=; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=0`}});
    }
    return new Response('Not found',{status:404,headers:security});
  }catch{
    // Do not log URLs, authorization codes, token responses, or cookie data.
    return page('Unable to connect. Check the Entra redirect URI, delegated permissions, admin consent, and server configuration, then try again.',503);
  }
}
