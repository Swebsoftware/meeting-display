export const states = ['available', 'busy', 'in_meeting', 'away', 'do_not_disturb', 'offline'] as const;
export type Status = typeof states[number];
export interface StatusResponse { status: Status; displayStatus: string; message: string; busyUntil: string | null; updatedAt: string; timeZone: string; demo: boolean }
export const labels: Record<Status, [string, string]> = {
  available: ['Available', 'Feel free to come in'], busy: ['Busy', 'Please do not disturb'],
  in_meeting: ['In a Meeting', 'Please come back later'], away: ['Away', 'Please check back soon'],
  do_not_disturb: ['Do Not Disturb', 'Please do not disturb'], offline: ['Status unavailable', 'Please check back soon'],
};
export function unavailable(): StatusResponse { return {status:'offline',displayStatus:labels.offline[0],message:labels.offline[1],busyUntil:null,updatedAt:new Date().toISOString(),timeZone:'UTC',demo:false}; }
export function formatTime(date: string | Date, timeZone: string) { return new Intl.DateTimeFormat('en-US', {hour:'numeric',minute:'2-digit',hour12:true,timeZone}).format(new Date(date)); }
export function retainStatus(last: StatusResponse | null, now = Date.now()): StatusResponse { return last && now - Date.parse(last.updatedAt) < 120_000 ? last : unavailable(); }
export function parseStatus(value: unknown): StatusResponse {
  const v = value as StatusResponse;
  if (!v || !states.includes(v.status) || typeof v.displayStatus !== 'string' || typeof v.message !== 'string' || typeof v.demo !== 'boolean' || !Number.isFinite(Date.parse(v.updatedAt)) || Date.parse(v.updatedAt) > Date.now()+60_000 || Date.now()-Date.parse(v.updatedAt)>120_000 || (v.busyUntil !== null && !Number.isFinite(Date.parse(v.busyUntil)))) throw new Error('Invalid status');
  formatTime(v.updatedAt, v.timeZone); return v;
}
