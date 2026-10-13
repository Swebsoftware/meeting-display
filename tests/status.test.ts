import {describe,it,expect,vi,afterEach} from 'vitest';
import {normalize} from '../worker/normalize';
import {formatTime,retainStatus,parseStatus} from '../src/status';
import worker from '../worker/index';
const now=Date.parse('2026-10-13T17:00:00Z');
const meeting={start:{dateTime:'2026-10-13T16:30:00'},end:{dateTime:'2026-10-13T18:30:00'},showAs:'busy'};
afterEach(()=>vi.useRealTimers());
describe('status priority',()=>{
 it('prioritizes presenting and DND over meetings',()=>{for(const activity of ['Presenting','DoNotDisturb'])expect(normalize({activity,availability:'Busy'},[meeting],'America/New_York',now).status).toBe('do_not_disturb');});
 it('maps calls and meetings without inventing end times',()=>{for(const activity of ['InACall','InAMeeting'])expect(normalize({activity,availability:'Busy'},[],'UTC',now)).toMatchObject({status:'in_meeting',busyUntil:null});});
 it('uses current busy calendar entries and UTC Graph times',()=>expect(normalize({activity:'Available',availability:'Available'},[meeting],'America/New_York',now)).toMatchObject({status:'in_meeting',busyUntil:'2026-10-13T18:30:00.000Z'}));
 it('ignores cancelled, all-day and free events',()=>{for(const extra of [{isCancelled:true},{isAllDay:true},{showAs:'free'}])expect(normalize({activity:'Available',availability:'Available'},[{...meeting,...extra}],'UTC',now).status).toBe('available');});
 it('maps all remaining states',()=>{for(const [availability,status] of [['Busy','busy'],['Away','away'],['BeRightBack','away'],['Available','available'],['Offline','offline'],['Unknown','offline']])expect(normalize({availability,activity:availability},[],'UTC',now).status).toBe(status);});
});
describe('reliability and times',()=>{
 it('formats Eastern daylight and standard time correctly',()=>{expect(formatTime('2026-10-13T18:30:00Z','America/New_York')).toBe('2:30 PM');expect(formatTime('2026-12-13T18:30:00Z','America/New_York')).toBe('1:30 PM');});
 it('retains recent status but expires it after two minutes',()=>{const data=normalize({activity:'Available',availability:'Available'},[],'UTC',now);expect(retainStatus(data,now+119999)).toBe(data);expect(retainStatus(data,now+120000).status).toBe('offline');});
 it('rejects invalid API data',()=>{vi.useFakeTimers();vi.setSystemTime(now);expect(()=>parseStatus({})).toThrow();const data=normalize({activity:'Available',availability:'Available'},[],'UTC',now);expect(parseStatus(data)).toEqual(data);expect(()=>parseStatus({...data,timeZone:'invalid'})).toThrow();});
});
describe('Worker API',()=>{
 const env={DEMO_MODE:'true',ASSETS:{fetch:async()=>new Response('asset')}};
 it('serves every demo state',async()=>{for(const status of ['available','busy','in_meeting','away','do_not_disturb','offline']){const r=await worker.fetch(new Request(`https://office.test/api/status?state=${status}`),env);expect(r.status).toBe(200);expect(await r.json()).toMatchObject({status,demo:true});expect(r.headers.get('Cache-Control')).toBe('no-store');}});
 it('rejects invalid state and method',async()=>{expect((await worker.fetch(new Request('https://office.test/api/status?state=invalid'),env)).status).toBe(400);expect((await worker.fetch(new Request('https://office.test/api/status',{method:'POST'}),env)).status).toBe(405);});
 it('returns a safe error when configuration is missing',async()=>{vi.spyOn(console,'error').mockImplementation(()=>{});const r=await worker.fetch(new Request('https://office.test/api/status'),{...env,DEMO_MODE:'false'});expect(r.status).toBe(503);expect(await r.json()).toEqual({error:'Status temporarily unavailable'});vi.restoreAllMocks();});
});
