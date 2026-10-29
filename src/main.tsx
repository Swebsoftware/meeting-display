import { Connect } from './Connect';
import React, {useEffect,useRef,useState} from 'react';
import {createRoot} from 'react-dom/client';
import {formatTime,parseStatus,retainStatus,states,type StatusResponse} from './status';
import './style.css';
function App(){
  const displayId=location.pathname.startsWith('/display/')?location.pathname.slice('/display/'.length):null;
  const [data,setData]=useState<StatusResponse|null>(null);
  const [connected,setConnected]=useState(true);
  const [reconnect,setReconnect]=useState(false);
  const [clock,setClock]=useState(new Date());
  const [demoState,setDemoState]=useState(new URLSearchParams(location.search).get('state')||'available');
  const last=useRef<StatusResponse|null>(null);
  const [awake,setAwake]=useState(false);
  useEffect(()=>{let dead=false;let inFlight=false;let controller:AbortController|undefined;
    async function poll(){if(dead || inFlight)return;inFlight=true;controller=new AbortController();const timeout=setTimeout(()=>controller?.abort(),12000);
      try{const response=await fetch(`/api/status?state=${encodeURIComponent(demoState)}${displayId?`&display=${encodeURIComponent(displayId)}`:''}`,{cache:'no-store',signal:controller.signal});if(!response.ok){const failure=await response.json().catch(()=>({}));if(!dead)setReconnect(failure.reconnectRequired===true);throw new Error('API unavailable');}const next=parseStatus(await response.json());if(!dead){last.current=next;setData(next);setConnected(true);setReconnect(false);}}
      catch{if(!dead){setData(retainStatus(last.current));setConnected(false);}}finally{clearTimeout(timeout);inFlight=false;}}
    void poll();const timer=setInterval(()=>void poll(),10000);
    const expire=setInterval(()=>{if(last.current && Date.now()-Date.parse(last.current.updatedAt)>=120000){setData(retainStatus(null));setConnected(false);}},1000);
    const resume=()=>{if(document.visibilityState==='visible')void poll();};document.addEventListener('visibilitychange',resume);
    return()=>{dead=true;controller?.abort();clearInterval(timer);clearInterval(expire);document.removeEventListener('visibilitychange',resume);};
  },[demoState,displayId]);
  useEffect(()=>{const timer=setInterval(()=>setClock(new Date()),1000);return()=>clearInterval(timer);},[]);
  useEffect(()=>{let lock:WakeLockSentinel|undefined;let disposed=false;
    const acquire=async()=>{try{if(document.visibilityState==='visible' && 'wakeLock' in navigator){const next=await navigator.wakeLock.request('screen');if(disposed){await next.release();return;}lock=next;setAwake(true);next.addEventListener('release',()=>setAwake(false));}}catch{setAwake(false);}};
    void acquire();document.addEventListener('visibilitychange',acquire);return()=>{disposed=true;document.removeEventListener('visibilitychange',acquire);void lock?.release();};
  },[]);
  const status=data?.status||'offline';const zone=data?.timeZone||Intl.DateTimeFormat().resolvedOptions().timeZone;
  return <main className={`sign ${status}`}>
    <header><div className="office"><span className="office-mark">◫</span><span className="office-label"><span className="office-title">{data?.jobTitle?.trim() || 'OFFICE'}</span>{data?.displayName?.trim() && <span className="office-name">{data.displayName.trim()}</span>}<span className="sub">AVAILABILITY DISPLAY</span></span></div><div className="date"><span>{new Intl.DateTimeFormat('en-US',{weekday:'long',month:'long',day:'numeric',timeZone:zone}).format(clock)}</span><strong>{formatTime(clock,zone)}</strong></div></header>
    <section className="display" aria-live="polite" aria-atomic="true"><div className="indicator" aria-hidden="true"><div className="halo"><div className="disc">{status==='available'?<svg viewBox="0 0 100 100"><path d="m25 51 17 17 34-37"/></svg>:status==='in_meeting'?<svg viewBox="0 0 100 100"><rect x="22" y="32" width="40" height="36" rx="8"/><path d="m63 43 17-10v34L63 57"/></svg>:status==='away'?<svg viewBox="0 0 100 100"><circle cx="50" cy="50" r="29"/><path d="M50 32v20l13 8"/></svg>:<svg viewBox="0 0 100 100"><path d="M29 50h42"/></svg>}</div></div><div className="status-tag"><i/>{status==='offline'?'UNAVAILABLE':status==='available'?'YOU’RE WELCOME':status==='away'?'BACK SOON':'QUIET, PLEASE'}</div></div>
    <div className="copy"><div className="eyebrow">{status==='available'?'THE DOOR IS OPEN':'CURRENT OFFICE STATUS'}</div><h1>{data?.displayStatus||'Status unavailable'}</h1><p>{data?.message||'Connecting to the office…'}</p>{data?.busyUntil && new Date(data.busyUntil)>clock && <div className="until"><span className="small-clock">◷</span> Busy until <strong>{formatTime(data.busyUntil,zone)}</strong></div>}</div></section>
    <footer><div className="connection"><i className={connected&&data?'live':''}/>{!data?'Connecting':!connected?'Reconnecting · status may be out of date':data.demo?'Demo display':'Live status'}<span className="separator">/</span><span>Updates automatically</span></div><button className="fullscreen" onClick={()=>{if(!document.fullscreenElement)void document.documentElement.requestFullscreen?.().catch(()=>{});else void document.exitFullscreen();}} aria-label="Toggle full screen">⛶ <span>Full screen</span></button></footer>
    <div className="powered-by">Powered by S. Web Software</div>
    {reconnect && <a className="reconnect-link" href="/connect">Reconnect Microsoft</a>}
    {data?.demo && <nav className="demo" aria-label="Demo states"><span>DEMO</span>{states.map(s=><button key={s} aria-pressed={demoState===s} onClick={()=>setDemoState(s)}>{s==='in_meeting'?'Meeting':s==='do_not_disturb'?'Do not disturb':s.charAt(0).toUpperCase()+s.slice(1)}</button>)}</nav>}
    <span className="sr-only">{awake?'Screen wake lock active':'Use tablet kiosk settings to keep the screen awake.'}</span>
  </main>;
}
if(location.pathname.startsWith('/display/')) {
  const id=location.pathname.slice('/display/'.length);
  document.querySelector('link[rel="manifest"]')?.setAttribute('href',`/api/manifest?display=${encodeURIComponent(id)}`);
}
createRoot(document.getElementById('root')!).render(<React.StrictMode>{(location.pathname.startsWith('/display/') || location.pathname==='/demo')?<App/>:<Connect/>}</React.StrictMode>);
if('serviceWorker' in navigator && import.meta.env.PROD) window.addEventListener('load',()=>{void navigator.serviceWorker.register('/sw.js').catch(()=>{});});
