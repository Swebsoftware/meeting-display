import {it,expect,vi} from 'vitest';
import worker from '../worker/index';
import {encrypt} from '../worker/auth';
import type {Env} from '../worker/graph';
it('isolates live status cache and token requests by personal display',async()=>{
 const storage=new Map<string,string>();
 const env:Env={AUTH_MODE:'delegated',APP_ORIGIN:'https://office.test',MICROSOFT_TENANT_ID:'isolation-tenant',MICROSOFT_CLIENT_ID:'isolation-client',MICROSOFT_CLIENT_SECRET:'test-only',TOKEN_ENCRYPTION_KEY:btoa('12345678901234567890123456789012'),CONNECTIONS:{get:async k=>storage.get(k)||null,put:async(k,v)=>{storage.set(k,v);}},ASSETS:{fetch:async()=>new Response('')}};
 const alice='C'.repeat(43),bob='D'.repeat(43);
 for(const id of [alice,bob])storage.set(`connection:${id}`,await encrypt({accessToken:id,refreshToken:'test-refresh',expiresAt:Date.now()+3600000,userId:id,clientId:env.MICROSOFT_CLIENT_ID,tenantId:env.MICROSOFT_TENANT_ID},env));
 const fetchMock=vi.fn(async(input:RequestInfo|URL,init?:RequestInit)=>{
  const url=String(input);const owner=(init!.headers as Record<string,string>).Authorization;
  if(url.endsWith('/presence'))return Response.json({availability:owner.endsWith(alice)?'Busy':'Available',activity:'Unknown'});
  if(url.includes('calendarView'))return Response.json({value:[]});
  return Response.json({timeZone:'America/New_York'});
 });vi.stubGlobal('fetch',fetchMock);
 try{
  const responses=await Promise.all([alice,bob].map(id=>worker.fetch(new Request(`https://office.test/api/status?display=${id}`),env)));
  expect(await responses[0].json()).toMatchObject({status:'busy'});expect(await responses[1].json()).toMatchObject({status:'available'});
  expect(await (await worker.fetch(new Request(`https://office.test/api/status?display=${alice}`),env)).json()).toMatchObject({status:'busy'});expect(fetchMock).toHaveBeenCalledTimes(6);
  expect((await worker.fetch(new Request('https://office.test/api/status?display=someone@example.com'),env)).status).toBe(404);
  const missing=await worker.fetch(new Request(`https://office.test/api/status?display=${'E'.repeat(43)}`),env);expect(missing.status).toBe(503);expect(await missing.json()).toMatchObject({reconnectRequired:true});
 }finally{vi.unstubAllGlobals();}
});

it('personal PWA installation starts on the same user display',async()=>{
 const id='H'.repeat(43);const response=await worker.fetch(new Request(`https://office.test/api/manifest?display=${id}`),{ASSETS:{fetch:async()=>new Response('')}});
 expect(response.headers.get('content-type')).toBe('application/manifest+json');expect(await response.json()).toMatchObject({id:`/display/${id}`,start_url:`/display/${id}`});
});
