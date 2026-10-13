import {test,expect} from '@playwright/test';
for(const viewport of [{width:1280,height:800},{width:768,height:1024},{width:1024,height:600}])test(`tablet layout ${viewport.width}x${viewport.height}`,async({page})=>{
 await page.setViewportSize(viewport);await page.goto('/');await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 for(const [button,heading] of [['Busy','Busy'],['Meeting','In a Meeting'],['Away','Away'],['Do not disturb','Do Not Disturb'],['Offline','Status unavailable'],['Available','Available']]){await page.getByRole('button',{name:button,exact:true}).click();await expect(page.getByRole('heading',{name:heading,exact:true})).toBeVisible();expect(await page.evaluate(()=>document.documentElement.scrollHeight<=innerHeight)).toBe(true);}
 await page.waitForTimeout(700);
 if(viewport.width===1280)await page.screenshot({path:'/tmp/office-display.png'});
});
test('polls and expires the last status on failures',async({page})=>{
 await page.clock.install();await page.goto('/');await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 await page.route('**/api/status*',route=>route.fulfill({status:503,json:{error:'unavailable'}}));
 await page.clock.fastForward(25000);await expect(page.getByText('Reconnecting · status may be out of date')).toBeVisible();await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
 await page.clock.fastForward(121000);await expect(page.getByRole('heading',{name:'Status unavailable',exact:true})).toBeVisible();
 await page.unroute('**/api/status*');await page.route('**/api/status*',async route=>{const response=await route.fetch();const data=await response.json();data.updatedAt=await page.evaluate(()=>new Date().toISOString());await route.fulfill({json:data});});await page.clock.fastForward(25000);await expect(page.getByRole('heading',{name:'Available',exact:true})).toBeVisible();
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
