import asyncio
from playwright.async_api import async_playwright
PW=open('/home/claude/pw.txt').read().strip()
async def main():
    async with async_playwright() as p:
        b=await p.chromium.launch();pg=await b.new_page(viewport={"width":1280,"height":900},accept_downloads=True);errs=[];pg.on("pageerror",lambda e:errs.append(str(e)))
        await pg.goto("file:///mnt/user-data/outputs/GIT_anak_toko.html");await pg.wait_for_selector("#login",state="visible",timeout=15000)
        print("login tampil:",True,"| upload tersembunyi:",not await pg.is_visible("#f"),"| tab terlihat sebelum masuk:",await pg.evaluate("document.querySelector('#out').style.display"))
        await pg.fill("#pw","salah-salah-123");await pg.click("#pwb");await pg.wait_for_timeout(1800);print("sandi salah ->",await pg.inner_text("#pwm"),"| out:",await pg.is_visible("#out"))
        await pg.fill("#pw",PW);await pg.click("#pwb");await pg.wait_for_selector("#res table",timeout=60000)
        print("header:",await pg.inner_text("header .hd>div>div"))
        print("tab:",await pg.evaluate("[...document.querySelectorAll('.tabs button')].filter(b=>b.offsetParent).map(b=>b.textContent)"))
        print("KPI:",(await pg.inner_text("#kpi")).replace("\n"," ")[:120])
        await pg.select_option("#bln","");await pg.wait_for_timeout(300)
        await pg.fill("#q","so1j");await pg.dispatch_event("#q","input");await pg.wait_for_timeout(300)
        print("profil:",(await pg.inner_text("#det")).replace("\n"," ")[:80],"| Share WA ada:",await pg.evaluate("!!document.querySelector('#sb')"))
        await pg.fill("#q","");await pg.dispatch_event("#q","input");await pg.wait_for_timeout(300)
        async with pg.expect_download() as d: await pg.click("#xa")
        dl=await d.value;await dl.save_as("/tmp/viewer.xlsx");print("unduh:",dl.suggested_filename,"| errors:",errs);await b.close()
asyncio.run(main())
