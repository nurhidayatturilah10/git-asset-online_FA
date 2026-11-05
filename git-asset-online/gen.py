import asyncio
from playwright.async_api import async_playwright
PW=open('/home/claude/pw.txt').read().strip()
async def main():
    async with async_playwright() as p:
        b=await p.chromium.launch();pg=await b.new_page(viewport={"width":1280,"height":900},accept_downloads=True);errs=[];pg.on("pageerror",lambda e:errs.append(str(e)))
        await pg.goto("file:///tmp/vt5/index.html?admin")
        await pg.set_input_files("#f","/mnt/user-data/uploads/REPORT_GIT.xlsx");await pg.wait_for_selector("#res table",timeout=120000)
        print("info:",await pg.inner_text("#info"),"| ref:",await pg.input_value("#ref"))
        await pg.evaluate("go('lk')");await pg.fill("#p1",PW);await pg.fill("#p2",PW)
        async with pg.expect_download(timeout=200000) as d: await pg.click("#mk")
        dl=await d.value;await dl.save_as("/mnt/user-data/outputs/data.js");print("data.js dibuat | errors:",errs);await b.close()
asyncio.run(main())
