import { useEffect, useState } from 'react';
interface Connection {connected:boolean;configured:boolean;demo:boolean;displayPath:string|null}
export function Connect() {
  const [connection,setConnection]=useState<Connection|null>(null);
  const [failed,setFailed]=useState(false);
  const [copied,setCopied]=useState(false);
  useEffect(()=>{const controller=new AbortController();fetch('/api/auth/connection',{cache:'no-store',signal:controller.signal}).then(async r=>{if(!r.ok)throw new Error();setConnection(await r.json());}).catch(()=>{if(!controller.signal.aborted)setFailed(true);});return()=>controller.abort();},[]);
  const displayUrl=connection?.displayPath ? new URL(connection.displayPath,location.origin).href : null;
  return <main className="connect-page"><div className="connect-card"><div className="eyebrow">OFFICE DISPLAY · CONNECTION</div><h1>Connect Microsoft</h1><p>Sign in with your work account. Get your own office display link, with your Teams availability and calendar busy-until time.</p><div className="connection-detail">{failed?'Unable to check connection. Please reload and try again.':!connection?'Checking connection…':!connection.configured?'Microsoft sign-in needs server configuration.':connection.connected?'Microsoft account connected.': 'Sign in to create your office display.'}</div>
    {displayUrl && <section className="personal-link"><label htmlFor="display-link">Your display link</label><input id="display-link" readOnly value={displayUrl} onFocus={e=>e.target.select()}/><div className="link-actions"><a className="connect-button" href={connection!.displayPath!}>Open my display</a><button className="copy-link" onClick={()=>{void navigator.clipboard.writeText(displayUrl).then(()=>setCopied(true)).catch(()=>setCopied(false));}}>{copied?'Copied':'Copy link'}</button></div><p className="connection-note">Open this link on your tablet. Anyone with the link can view your availability; share it only with people you trust.</p></section>}
    {connection?.demo && <p className="connection-note">Demo mode is enabled; the display will use sample availability.</p>}
    <form method="post" action="/api/auth/start"><button disabled={!connection?.configured} className="connect-button">{connection?.connected?'Reconnect Microsoft':'Connect Microsoft'}</button></form>
    {connection?.connected && <form method="post" action="/api/auth/logout"><button className="sign-out">Sign out of settings</button></form>}
    <p className="connection-note">For Microsoft 365 accounts in our organization. Each account has its own display. Tokens stay encrypted on the server; meeting details are never displayed. Your organization may require consent or administrator approval.</p></div></main>;
}
