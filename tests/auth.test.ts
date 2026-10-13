import {beforeEach,afterEach,describe,it,expect,vi} from 'vitest';
import {authRoute,encrypt,decrypt,delegatedToken,ReconnectRequired} from '../worker/auth';
import type {Env} from '../worker/graph';
let storage: Map<string,string>;
let env:Env;
beforeEach(()=>{storage=new Map();env={AUTH_MODE:'delegated',APP_ORIGIN:'https://office.test',MICROSOFT_TENANT_ID:'tenant-test',MICROSOFT_CLIENT_ID:'client-test',MICROSOFT_CLIENT_SECRET:'test-secret',TARGET_USER_ID:'sam@ontrackteam.co',TOKEN_ENCRYPTION_KEY:btoa('12345678901234567890123456789012'),CONNECTIONS:{get:async key=>storage.get(key)||null,put:async(key,value)=>{storage.set(key,value);}},ASSETS:{fetch:async()=>new Response('')}};});
afterEach(()=>{vi.unstubAllGlobals();vi.restoreAllMocks();});
async function start(){
 const response=await authRoute(new Request('https://office.test/api/auth/start',{method:'POST',headers:{Origin:'https://office.test'}}),env);
 const cookie=response.headers.get('set-cookie')!.split(';')[0];
 const state=new URL(response.headers.get('Location')!).searchParams.get('state')!;
 return {response,cookie,state};
}
function callback(state:string,cookie:string){return new Request(`https://office.test/api/auth/callback?code=test-code&state=${state}`,{headers:{Cookie:cookie}});}
function graphMock(email='sam@ontrackteam.co'){
 return vi.fn(async(input:RequestInfo|URL,init?:RequestInit)=>{
  if(String(input).includes('/token'))return Response.json({access_token:'test-access',refresh_token:'test-refresh',expires_in:3600});
  return Response.json({id:'user-id',userPrincipalName:email,mail:email});
 });
}
describe('Microsoft connection security',()=>{
 it('encrypts and authenticates stored tokens',async()=>{const encrypted=await encrypt({secret:'not-plaintext'},env);expect(encrypted).not.toContain('not-plaintext');expect(await decrypt(encrypted,env)).toEqual({secret:'not-plaintext'});await expect(decrypt(encrypted.slice(0,-4)+'xxxx',env)).rejects.toThrow();});
 it('uses an HTTPS fixed callback, state, PKCE and a protected cookie',async()=>{const {response}=await start();expect(response.status).toBe(303);const url=new URL(response.headers.get('location')!);expect(url.searchParams.get('redirect_uri')).toBe('https://office.test/api/auth/callback');expect(url.searchParams.get('code_challenge_method')).toBe('S256');expect(url.searchParams.get('scope')).toContain('offline_access');expect(response.headers.get('set-cookie')).toContain('HttpOnly; Secure; SameSite=Lax');});
 it('rejects GET initiation and cross-origin POSTs',async()=>{expect((await authRoute(new Request('https://office.test/api/auth/start'),env)).status).toBe(405);expect((await authRoute(new Request('https://office.test/api/auth/start',{method:'POST',headers:{Origin:'https://attacker.test'}}),env)).status).toBe(403);});
 it('rejects missing cookies and mismatched state before exchanging codes',async()=>{const fetchMock=vi.fn();vi.stubGlobal('fetch',fetchMock);expect((await authRoute(callback('bad',''),env)).status).toBe(400);const {cookie}=await start();expect((await authRoute(callback('bad',cookie),env)).status).toBe(400);expect(fetchMock).not.toHaveBeenCalled();});
 it('rejects expired callbacks',async()=>{const cookie=`__Host-office-oauth=${await encrypt({state:'state',verifier:'verifier',expiresAt:Date.now()-1},env)}`;expect((await authRoute(callback('state',cookie),env)).status).toBe(400);});
 it('rejects another account without replacing an existing connection',async()=>{storage.set('microsoft-connection-v1','existing');vi.stubGlobal('fetch',graphMock('other@example.com'));const {cookie,state}=await start();expect((await authRoute(callback(state,cookie),env)).status).toBe(403);expect(storage.get('microsoft-connection-v1')).toBe('existing');});
 it('stores only the authorized account encrypted and reuses its token',async()=>{const mock=graphMock();vi.stubGlobal('fetch',mock);const {cookie,state}=await start();const r=await authRoute(callback(state,cookie),env);expect(r.status).toBe(303);expect(r.headers.get('location')).toBe('/connect?connected=1');expect(storage.get('microsoft-connection-v1')).not.toContain('test-refresh');expect(await delegatedToken(env)).toBe('test-access');expect(mock).toHaveBeenCalledTimes(2);const status=await authRoute(new Request('https://office.test/api/auth/connection'),env);expect(await status.json()).toMatchObject({connected:true});});
 it('refreshes expired tokens and persists rotated refresh tokens',async()=>{await storage.set('microsoft-connection-v1',await encrypt({accessToken:'old',refreshToken:'old-refresh',expiresAt:0,target:env.TARGET_USER_ID,clientId:env.MICROSOFT_CLIENT_ID,tenantId:env.MICROSOFT_TENANT_ID,userId:'user-id'},env));const mock=graphMock();vi.stubGlobal('fetch',mock);expect(await delegatedToken(env)).toBe('test-access');const body=mock.mock.calls[0][1]?.body as URLSearchParams;expect(body.get('grant_type')).toBe('refresh_token');expect(body.get('refresh_token')).toBe('old-refresh');expect(await decrypt(storage.get('microsoft-connection-v1')!,env)).toMatchObject({refreshToken:'test-refresh'});});
 it('signals reconnect on revoked authorization',async()=>{storage.set('microsoft-connection-v1',await encrypt({accessToken:'old',refreshToken:'old',expiresAt:0,target:env.TARGET_USER_ID,clientId:env.MICROSOFT_CLIENT_ID,tenantId:env.MICROSOFT_TENANT_ID},env));vi.stubGlobal('fetch',async()=>Response.json({error:'invalid_grant'},{status:400}));await expect(delegatedToken(env)).rejects.toBeInstanceOf(ReconnectRequired);});
 it('fails closed when target or app configuration changes',async()=>{storage.set('microsoft-connection-v1',await encrypt({target:'someone-else'},env));await expect(delegatedToken(env)).rejects.toBeInstanceOf(ReconnectRequired);});
});
