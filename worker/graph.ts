import { delegatedToken, ReconnectRequired } from './auth';
import { findIana } from 'windows-iana';
import { normalize, type CalendarEvent, type Presence } from './normalize';
export interface Env { AUTH_MODE?:string; APP_ORIGIN?:string; TOKEN_ENCRYPTION_KEY?:string; CONNECTIONS?:{get(key:string):Promise<string|null>;put(key:string,value:string):Promise<void>}; DEMO_MODE?:string; DISPLAY_TIME_ZONE?:string; MICROSOFT_TENANT_ID?:string; MICROSOFT_CLIENT_ID?:string; MICROSOFT_CLIENT_SECRET?:string; TARGET_USER_ID?:string; ASSETS:{fetch(request:Request):Promise<Response>} }
let token: {value:string;expires:number;key:string} | undefined;
let pendingToken: Promise<string> | undefined;
async function accessToken(env:Env):Promise<string> {
  const key=`${env.MICROSOFT_TENANT_ID}:${env.MICROSOFT_CLIENT_ID}`;
  if(token?.key===key && token.expires>Date.now()+60_000) return token.value;
  if(pendingToken) return pendingToken;
  pendingToken=(async()=>{
    const response=await fetch(`https://login.microsoftonline.com/${encodeURIComponent(env.MICROSOFT_TENANT_ID!)}/oauth2/v2.0/token`,{method:'POST',body:new URLSearchParams({client_id:env.MICROSOFT_CLIENT_ID!,client_secret:env.MICROSOFT_CLIENT_SECRET!,grant_type:'client_credentials',scope:'https://graph.microsoft.com/.default'}),signal:AbortSignal.timeout(10_000)});
    if(!response.ok) throw new Error(`Microsoft token request failed (${response.status})`);
    const body=await response.json() as {access_token:string;expires_in:number};
    if(!body.access_token || !body.expires_in) throw new Error('Invalid token response');
    token={value:body.access_token,expires:Date.now()+body.expires_in*1000,key}; return token.value;
  })();
  try { return await pendingToken; } finally { pendingToken=undefined; }
}
export async function microsoftStatus(env:Env) {
  for(const name of ['MICROSOFT_TENANT_ID','MICROSOFT_CLIENT_ID','MICROSOFT_CLIENT_SECRET','TARGET_USER_ID'] as const) if(!env[name]) throw new Error(`Missing ${name}`);
  const delegated=env.AUTH_MODE==='delegated';
  const bearer=delegated?await delegatedToken(env):await accessToken(env);
  async function get<T>(path:string):Promise<T> {
    const response=await fetch(`https://graph.microsoft.com/v1.0/${delegated?'me':`users/${encodeURIComponent(env.TARGET_USER_ID!)}`}/${path}`,{headers:{Authorization:`Bearer ${bearer}`,Prefer:'outlook.timezone="UTC"'},signal:AbortSignal.timeout(10_000)});
    if(!response.ok) {if(response.status===401) {token=undefined;if(delegated)throw new ReconnectRequired();}throw new Error(`Graph request failed (${response.status})`);}
    return await response.json() as T;
  }
  const now=Date.now();
  const query=new URLSearchParams({startDateTime:new Date(now-86400_000).toISOString(),endDateTime:new Date(now+86400_000).toISOString(),'$select':'start,end,showAs,isCancelled,isAllDay','$top':'100'});
  const [presence,calendar,settings]=await Promise.all([get<Presence>('presence'),get<{value:CalendarEvent[];'@odata.nextLink'?:string}>(`calendarView?${query}`),get<{timeZone:string}>('mailboxSettings?$select=timeZone')]);
  // Reject incomplete pages rather than silently omitting an active event.
  if(calendar['@odata.nextLink']) throw new Error('Calendar page limit exceeded');
  let timeZone=env.DISPLAY_TIME_ZONE || settings.timeZone;
  try { new Intl.DateTimeFormat('en-US',{timeZone}); } catch { timeZone=findIana(timeZone)[0]; }
  if(!timeZone) throw new Error('Unsupported mailbox time zone');
  return normalize(presence,calendar.value,timeZone,now);
}
