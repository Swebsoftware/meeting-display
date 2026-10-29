import {test,expect} from '@playwright/test';
// Optional separate demo Worker avoids changing the production live configuration.
if(process.env.TEST_API_ORIGIN)test.beforeEach(async({page})=>{await page.route('**/api/status*',async route=>{const original=new URL(route.request().url());const response=await route.fetch({url:`${process.env.TEST_API_ORIGIN}${original.pathname}${original.search}`});await route.fulfill({response});});});
for(const viewport of [{width:1280,height:800},{width:768,height:1024},{width:1024,height:600}])test(`tablet layout ${viewport.width}x${viewport.height}`,async({page})=>{
 await page.setViewportSize(viewport);await page.goto('/');await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 for(const [button,heading] of [['Busy','Busy'],['Meeting','In a Meeting'],['Away','Away'],['Do not disturb','Do Not Disturb'],['Offline','Status unavailable'],['Available','Available']]){await page.getByRole('button',{name:button,exact:true}).click();await expect(page.getByRole('heading',{name:heading,exact:true})).toBeVisible();expect(await page.evaluate(()=>document.documentElement.scrollHeight<=innerHeight)).toBe(true);}
 await page.waitForTimeout(700);
 if(viewport.width===1280)await page.screenshot({path:'/tmp/office-display.png'});
});
test('polls and expires the last status on failures',async({page})=>{
 await page.clock.install();await page.goto('/');await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 const failRoute=(route:import('@playwright/test').Route)=>route.fulfill({status:503,json:{error:'unavailable'}});await page.route('**/api/status*',failRoute);
 await page.clock.fastForward(10000);await expect(page.getByText('Reconnecting · status may be out of date')).toBeVisible();await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 await page.clock.fastForward(121000);await expect(page.getByRole('heading',{name:'Status unavailable',exact:true})).toBeVisible();
 await page.unroute('**/api/status*',failRoute);await page.route('**/api/status*',async route=>{const requestUrl=new URL(route.request().url());const response=await route.fetch(process.env.TEST_API_ORIGIN?{url:`${process.env.TEST_API_ORIGIN}${requestUrl.pathname}${requestUrl.search}`} : undefined);const data=await response.json();data.updatedAt=await page.evaluate(()=>new Date().toISOString());await route.fulfill({json:data});});await page.clock.fastForward(10000);await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
});
test('connection page exposes a sign-in action without changing the door layout',async({page})=>{
 await page.route('**/api/auth/connection',route=>route.fulfill({json:{connected:false,configured:true,demo:true}}));
 await page.goto('/connect');await expect(page.getByRole('heading',{name:'Connect Microsoft'})).toBeVisible();await expect(page.getByRole('button',{name:'Connect Microsoft'})).toBeEnabled();
 await page.route('**/api/auth/start',route=>route.fulfill({status:200,contentType:'text/html',body:'<h1>Microsoft sign-in request</h1>'}));await page.getByRole('button',{name:'Connect Microsoft'}).click();await expect(page.getByRole('heading',{name:'Microsoft sign-in request'})).toBeVisible();
});
test('production referrer policy preserves the same-origin sign-in POST',async({page})=>{
 // Use the shipped header policy, rather than Vite's default headers.
 const {readFileSync}=await import('node:fs');
 const policy=readFileSync('public/_headers','utf8').match(/Referrer-Policy:\s*(\S+)/)![1];
 await page.route('https://office.test/connect',route=>route.fulfill({contentType:'text/html',headers:{'Referrer-Policy':policy},body:'<form method="post" action="/api/auth/start"><button>Connect Microsoft</button></form>'}));
 let origin:string|undefined;
 await page.route('https://office.test/api/auth/start',route=>{origin=route.request().headers().origin;return route.fulfill({body:'Sign-in started'});});
 await page.goto('https://office.test/connect');await page.getByRole('button',{name:'Connect Microsoft'}).click();await page.waitForURL('**/api/auth/start');expect(origin).toBe('https://office.test');
});
test('each signed-in user gets a separate personal link and display request',async({browser})=>{
 for(const [id,status,label] of [['F'.repeat(43),'available','Available'],['G'.repeat(43),'busy','Busy']]){
  const context=await browser.newContext();const page=await context.newPage();
  await page.route('**/api/auth/connection',route=>route.fulfill({json:{connected:true,configured:true,demo:false,displayPath:`/display/${id}`}}));
  let requestedDisplay:string|null=null;
  await page.route('**/api/status*',route=>{requestedDisplay=new URL(route.request().url()).searchParams.get('display');return route.fulfill({json:{status,displayStatus:label,message:'Test account status',busyUntil:null,updatedAt:new Date().toISOString(),timeZone:'America/New_York',demo:false}});});
  await page.goto('/connect');await expect(page.getByLabel('Your display link')).toHaveValue(`http://127.0.0.1:5173/display/${id}`);
  await page.getByRole('link',{name:'Open my display'}).click();await expect(page.getByRole('heading',{name:label,exact:true})).toBeVisible();expect(requestedDisplay).toBe(id);await context.close();
 }
});
for(const viewport of [{width:960,height:540},{width:1024,height:600},{width:800,height:1280}])test(`legacy tablet keeps round indicators at ${viewport.width}x${viewport.height}`,async({page})=>{
 await page.setViewportSize(viewport);
 await page.route('**/api/status*',route=>route.fulfill({json:{status:'available',displayStatus:'Available',message:'Feel free to come in',busyUntil:null,updatedAt:new Date().toISOString(),timeZone:'America/New_York',demo:false}}));
 await page.goto('/');await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 // Simulate missing aspect-ratio and small viewport unit support on older Android browsers.
 await page.addStyleTag({content:'.halo{aspect-ratio:auto!important}.sign{height:100vh!important;min-height:100vh!important}'});
 for(const selector of ['.halo','.disc','.disc svg']){const bounds=await page.locator(selector).boundingBox();expect(bounds).not.toBeNull();expect(Math.abs(bounds!.width-bounds!.height)).toBeLessThan(2);}
 const sign=await page.locator('.sign').boundingBox();expect(sign!.height).toBe(viewport.height);
 expect(await page.evaluate(()=>document.documentElement.scrollHeight<=innerHeight && document.documentElement.scrollWidth<=innerWidth)).toBe(true);
 if(viewport.width===1024)await page.screenshot({path:'/tmp/office-tablet-fixed.png'});
});
test('shows Microsoft job title and handles missing titles on tablets',async({page})=>{
 await page.setViewportSize({width:1024,height:600});
 for(const title of ['Senior Accounting and Operations Manager',null]){
  await page.route('**/api/status*',route=>route.fulfill({json:{status:'available',displayStatus:'Available',message:'Feel free to come in',busyUntil:null,updatedAt:new Date().toISOString(),timeZone:'America/New_York',demo:false,jobTitle:title}}));
  await page.goto('/');await expect(page.locator('.office-title')).toHaveText(title || 'OFFICE');
  const office=await page.locator('.office').boundingBox();const date=await page.locator('.date').boundingBox();expect(office!.x+office!.width).toBeLessThan(date!.x);
  expect(await page.evaluate(()=>document.documentElement.scrollHeight<=innerHeight)).toBe(true);
 }
});
