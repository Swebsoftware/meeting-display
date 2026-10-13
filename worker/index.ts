import { authRoute, ReconnectRequired } from './auth';
import { microsoftStatus, type Env } from './graph';
import { labels, states, type StatusResponse } from '../src/status';
let cached: {data:StatusResponse;expires:number} | undefined;
let pending: Promise<StatusResponse> | undefined;
export default { async fetch(request:Request,env:Env):Promise<Response> {
  const url=new URL(request.url);
  if(url.pathname.startsWith('/api/auth/')) return authRoute(request,env);
  if(!url.pathname.startsWith('/api/')) return env.ASSETS.fetch(request);
  const headers={'Cache-Control':'no-store','Content-Type':'application/json','X-Content-Type-Options':'nosniff'};
  if(url.pathname!=='/api/status') return Response.json({error:'Not found'},{status:404,headers});
  if(request.method!=='GET') return Response.json({error:'Method not allowed'},{status:405,headers:{...headers,Allow:'GET'}});
  try {
    if(env.DEMO_MODE==='true') {
      const requested=url.searchParams.get('state') || 'available';
      if(!states.includes(requested as typeof states[number])) return Response.json({error:'Invalid demo state'},{status:400,headers});
      const status=requested as typeof states[number];
      return Response.json({status,displayStatus:labels[status][0],message:labels[status][1],busyUntil:status==='in_meeting'?new Date(Date.now()+25*60_000).toISOString():null,updatedAt:new Date().toISOString(),timeZone:env.DISPLAY_TIME_ZONE || 'America/New_York',demo:true},{headers});
    }
    if(!cached || cached.expires<=Date.now()) {
      pending ??= microsoftStatus(env);
      try {cached={data:await pending,expires:Date.now()+20_000};} finally {pending=undefined;}
    }
    return Response.json(cached.data,{headers});
  } catch(error) {
    // Never log raw Graph responses or credentials.
    console.error(error instanceof Error?error.message:'Status service failed');
    return Response.json({error:'Status temporarily unavailable',reconnectRequired:error instanceof ReconnectRequired},{status:503,headers:{...headers,'Retry-After':'25'}});
  }
}};
