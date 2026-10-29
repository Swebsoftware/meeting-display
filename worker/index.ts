import { authRoute, ReconnectRequired, validDisplayId } from './auth';
import { microsoftStatus, type Env } from './graph';
import { labels, states, type StatusResponse } from '../src/status';
const cache=new Map<string,{data:StatusResponse;expires:number}>();
const pending=new Map<string,Promise<StatusResponse>>();
export default { async fetch(request:Request,env:Env):Promise<Response> {
  const url=new URL(request.url);
  if(url.pathname.startsWith('/api/auth/')) return authRoute(request,env);
  if(!url.pathname.startsWith('/api/')) return env.ASSETS.fetch(request);
  const headers={'Cache-Control':'no-store','Content-Type':'application/json','X-Content-Type-Options':'nosniff'};
  if(url.pathname==='/api/manifest') {
    const id=url.searchParams.get('display');
    if(request.method!=='GET' || !id || !validDisplayId(id)) return Response.json({error:'Invalid display'},{status:400,headers});
    return Response.json({id:`/display/${id}`,name:'Office Availability',short_name:'Office',start_url:`/display/${id}`,scope:'/',display:'standalone',background_color:'#101915',theme_color:'#101915',icons:[192,512].map(size=>({src:`/icon-${size}.png`,sizes:`${size}x${size}`,type:'image/png',purpose:'any maskable'}))},{headers:{...headers,'Content-Type':'application/manifest+json'}});
  }
  if(url.pathname!=='/api/status') return Response.json({error:'Not found'},{status:404,headers});
  if(request.method!=='GET') return Response.json({error:'Method not allowed'},{status:405,headers:{...headers,Allow:'GET'}});
  try {
    if(env.DEMO_MODE==='true') {
      const requested=url.searchParams.get('state') || 'available';
      if(!states.includes(requested as typeof states[number])) return Response.json({error:'Invalid demo state'},{status:400,headers});
      const status=requested as typeof states[number];
      return Response.json({status,displayStatus:labels[status][0],message:labels[status][1],busyUntil:status==='in_meeting'?new Date(Date.now()+25*60_000).toISOString():null,updatedAt:new Date().toISOString(),timeZone:env.DISPLAY_TIME_ZONE || 'America/New_York',demo:true},{headers});
    }
    const requestedDisplay=url.searchParams.get('display');
    if(!requestedDisplay || !validDisplayId(requestedDisplay)) return Response.json({error:'Display not found'},{status:404,headers});
    const displayId=requestedDisplay;
    const cacheKey=`${env.MICROSOFT_TENANT_ID}:${env.MICROSOFT_CLIENT_ID}:${displayId || 'legacy'}`;
    let cached=cache.get(cacheKey);
    if(!cached || cached.expires<=Date.now()) {
      let operation=pending.get(cacheKey);
      if(!operation){operation=microsoftStatus(env,displayId);pending.set(cacheKey,operation);}
      try {
        cached={data:await operation,expires:Date.now()+5_000};
        if(cache.size>=256)cache.delete(cache.keys().next().value!);
        cache.set(cacheKey,cached);
      } finally {if(pending.get(cacheKey)===operation)pending.delete(cacheKey);}
    }
    return Response.json(cached.data,{headers});
  } catch(error) {
    // Never log raw Graph responses or credentials.
    console.error(error instanceof Error?error.message:'Status service failed');
    return Response.json({error:'Status temporarily unavailable',reconnectRequired:error instanceof ReconnectRequired},{status:503,headers:{...headers,'Retry-After':'25'}});
  }
}};
