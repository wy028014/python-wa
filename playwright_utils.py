import asyncio
import subprocess
from playwright.async_api import async_playwright

class PW:
    def __init__(self):
        self.browser = None
        self.context = None

    async def init(self):
        playwright = await async_playwright().start()
        self.browser = await playwright.chromium.launch(headless=False)
        self.context = await self.browser.new_context()
        page = await self.context.new_page()
        flag_login = await page.is_visible('xpath=//*[@id="loginBtn"]')
        while not flag_login:
            try:
                await page.goto('https://10.3.2.201:9943/rntibp/login.html')
            except:
                await page.click('#details-button')
                subprocess.run(r'C:\Personal\Projects\cyber\src\utils\index.exe')
                await page.click('#proceed-link')
            flag_login = await page.is_visible('xpath=//*[@id="loginBtn"]')
        await page.wait_for_timeout(2000)
        await page.click('xpath=//*[@id="loginBtn"]')
        await page.wait_for_timeout(2000)
        result = await page.is_visible('xpath=/html/body/div[2]/div[2]/div[2]/div/div[1]/a[1]')
        return result

    async def gjcx(self, date_start, date_end, id_no):
        page = await self.context.new_page()
        flag_login = await page.is_visible('xpath=//*[@id="loginBtn"]')
        while not flag_login:
            try:
                await page.goto('https://10.3.2.201:9943/rntibp/login.html')
            except:
                await page.click('#details-button')
                subprocess.run(r'C:\Personal\Projects\cyber\src\utils\index.exe')
                await page.click('#proceed-link')
            flag_login = await page.is_visible('xpath=//*[@id="loginBtn"]')
        await page.wait_for_timeout(2000)
        await page.click('xpath=//*[@id="loginBtn"]')
        await page.wait_for_timeout(2000)
        await page.goto('https://10.3.2.201:9943/rntibp/view/complex/trackQuery.html')
        await page.evaluate(f'''
        document.getElementsByClassName("main-padding")[0].style.backgroundColor="#000000";
        document.getElementsByName("startDate")[0].removeAttribute("readonly");
        document.getElementsByName("endDate")[0].removeAttribute("readonly");
        document.getElementsByClassName("dhxform_control")[0].id="startDate";
        document.getElementsByClassName("dhxform_control")[1].id="endDate";
        document.getElementsByClassName("dhxform_control")[3].id="idNo";
        document.getElementsByClassName("dhxform_btn")[1].id="download";
        document.getElementsByName("startDate")[0].value="{date_start}";
        document.getElementsByName("endDate")[0].value="{date_end}";
        ''')
        await page.fill('xpath=/html/body/div[1]/div/div/div/div[1]/div[2]/div/div[4]/div/div[2]/input', id_no)
        await page.click('#startDate')
        await page.click('#endDate')
        await page.click('#idNo')
        async with page.expect_download() as download_info:
            await page.click('#download')
        download = await download_info.value
        await download.save_as(r'C:\Personal\Projects\Data\轨迹查询.xlsx')
        await page.close()
        # 这里需要调用 exceljs_utils 中的读取函数
        from exceljs_utils import exceljs
        return await asyncio.to_thread(exceljs.read_file, r'C:\Personal\Projects\Data\轨迹查询.xlsx')

    # 其他方法类似实现，如 zzcx、gjcx_pl、zzcx_pl 等