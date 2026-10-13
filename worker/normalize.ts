import { labels, type Status, type StatusResponse } from '../src/status';
export interface Presence { availability: string; activity: string }
export interface CalendarEvent { start: {dateTime:string}; end:{dateTime:string}; showAs:string; isCancelled?:boolean; isAllDay?:boolean }
export function graphDate(date: string) { return Date.parse(/(?:Z|[+-]\d\d:\d\d)$/.test(date) ? date : `${date}Z`); }
export function normalize(p: Presence, events: CalendarEvent[], timeZone: string, now = Date.now()): StatusResponse {
  const current = events.filter(e => !e.isCancelled && !e.isAllDay && ['busy','oof'].includes(e.showAs) && graphDate(e.start.dateTime)<=now && graphDate(e.end.dateTime)>now);
  let status:Status;
  if (['Presenting','DoNotDisturb'].includes(p.activity) || p.availability === 'DoNotDisturb') status='do_not_disturb';
  else if (['InACall','InAMeeting'].includes(p.activity) || current.length) status='in_meeting';
  else if (p.availability === 'Busy') status='busy';
  else if (['Away','BeRightBack'].includes(p.availability)) status='away';
  else if (p.availability === 'Available') status='available';
  else status='offline';
  const end = current.length ? Math.max(...current.map(e => graphDate(e.end.dateTime))) : null;
  return {status,displayStatus:labels[status][0],message:labels[status][1],busyUntil:status==='in_meeting' && end ? new Date(end).toISOString():null,updatedAt:new Date(now).toISOString(),timeZone,demo:false};
}
