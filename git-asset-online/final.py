import asyncio
from playwright.async_api import async_playwright
PW=open('/home/claude/pw.txt').read().strip();IP="192.0.2.2"
async def run(b,url):
    pg=await b.new_page(viewport={"width":1100,"height":800},accept_downloads=True);errs=[];pg.on("pageerror",lambda e:errs.append(str(e)))
    await pg.goto(url);await pg.wait_for_selector("#login",state="visible",timeout=15000)
    await pg.fill("#pw",PW);await pg.click("#pwb");await pg.wait_for_selector("#res table",timeout=60000)
    await pg.select_option("#bln","");await pg.wait_for_timeout(300)
    async with pg.expect_download() as d: await pg.click("#xa")
    dl=await d.value;await dl.save_as("/tmp/final.xlsx")
    from openpyxl import load_workbook
    wb=load_workbook("/tmp/final.xlsx");print(url.split("//")[0],url.split("/")[2] if "http" in url else "file","| aman:",await pg.evaluate("isSecureContext"),"| buka OK | Excel:",{w.title:w.max_row-1 for w in wb},"| errors:",errs);await pg.close()
async def main():
    async with async_playwright() as p:
        b=await p.chromium.launch()
        await run(b,"file:///mnt/user-data/outputs/GIT_anak_toko.html")
        await run(b,f"http://{IP}:8125/final_viewer.html")
        await run(b,"http://localhost:8125/final_viewer.html");await b.close()
asyncio.run(main())
