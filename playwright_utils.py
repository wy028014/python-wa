import asyncio
import pyautogui
import pygetwindow as gw
from playwright.async_api import async_playwright
from exceljs_utils import exceljs
import os


class PW:
    def __init__(self):
        self.browser = None
        self.context = None
        self.timeout = 30000  # 30 秒超时时间
        self.dir = r'C:\Personal\Projects\Python\python-wa\data'
        os.makedirs(os.path.dirname(self.dir), exist_ok=True)

    async def init(self):
        if not self.browser or await self.browser.is_closed():
            playwright = await async_playwright().start()
            self.browser = await playwright.chromium.launch(
                headless=False,
                args=['--start-maximized'],
                ignore_https_errors=True
            )
            self.context = await self.browser.new_context()
        page = await self.context.new_page()
        try:
            await page.goto('https://10.3.2.201:9943/rntibp/login.html', timeout=self.timeout)
        except Exception as e:
            try:
                await page.click('#details-button', timeout=self.timeout)
                await page.click('#proceed-link', timeout=self.timeout)
                window_name = "Your Actual Window Name"  # 修改为实际的窗口名称
                for window in gw.getAllWindows():
                    if window.title == window_name and window.isVisible and window.isActive:
                        pyautogui.typewrite('111111')
                        pyautogui.press('enter')
                        break
            except Exception as inner_e:
                print(f"Error handling security page: {inner_e}")
        flag_login = await page.is_visible('xpath=//*[@id="loginBtn"]', timeout=self.timeout)
        if flag_login:
            await page.click('xpath=//*[@id="loginBtn"]', timeout=self.timeout)
            await page.wait_for_timeout(2000)
        result = await page.is_visible('xpath=/html/body/div[2]/div[2]/div[2]/div/div[1]/a[1]', timeout=self.timeout)
        return result

    async def _fill_form_and_download(self, page, form_script, fill_actions, save_path):
        await page.evaluate(form_script)
        for selector, value in fill_actions:
            await page.fill(selector, value)
        for selector in ['#startDate', '#endDate', '#idNo', '#trainDate', '#boardTrainCode', '#fromStation', '#toStation']:
            if await page.is_visible(selector):
                await page.click(selector)
        try:
            async with page.expect_download(timeout=self.timeout) as download_info:
                await page.click('#download')
            download = await download_info.value
            await download.save_as(save_path)
            result = await asyncio.to_thread(exceljs.read_file, save_path)
            return result
        except asyncio.TimeoutError:
            print(f"Download timed out after {self.timeout / 1000} seconds.")
            return []

    async def _perform_query(self, page, query_type, params):
        all_results = []
        if query_type == 'glcx':
            for date_start, date_end, id_no in params:
                form_script = f'''
                document.getElementsByClassName("main-padding")[0].style.backgroundColor="#000000";
                document.getElementsByName("startDate")[0].removeAttribute("readonly");
                document.getElementsByName("endDate")[0].removeAttribute("readonly");
                document.getElementsByClassName("dhxform_control")[0].id="startDate";
                document.getElementsByClassName("dhxform_control")[1].id="endDate";
                document.getElementsByClassName("dhxform_control")[3].id="idNo";
                document.getElementsByClassName("dhxform_btn")[1].id="download";
                document.getElementsByName("startDate")[0].value="{date_start}";
                document.getElementsByName("endDate")[0].value="{date_end}";
                '''
                fill_actions = [
                    ('xpath=/html/body/div[1]/div/div/div/div[1]/div[2]/div/div[4]/div/div[2]/input', id_no)
                ]
                save_path = os.path.join(self.dir, r'关联查询.xlsx')
                result = await self._fill_form_and_download(page, form_script, fill_actions, save_path)
                all_results.extend(result)
        elif query_type == 'zzcx':
            for date_start, train_code, from_station, to_station in params:
                form_script = f'''
                document.getElementsByClassName("main-padding")[0].style.backgroundColor="#000000";
                document.getElementsByName("trainDate")[0].removeAttribute("readonly");
                document.getElementsByClassName("dhxform_control")[0].id="trainDate";
                document.getElementsByClassName("dhxform_control")[1].id="boardTrainCode";
                document.getElementsByClassName("dhxform_control")[2].id="fromStation";
                document.getElementsByClassName("dhxform_control")[3].id="toStation";
                document.getElementsByClassName("dhxform_btn")[1].id="download";
                document.getElementsByName("trainDate")[0].value="{date_start}";
                document.getElementsByName("boardTrainCode")[0].value="{train_code}";
                document.getElementsByName("fromStation")[0].value="{from_station}";
                document.getElementsByName("toStation")[0].value="{to_station}";
                '''
                fill_actions = [
                    ('xpath=/html/body/div[1]/div/div/div/div[1]/div[2]/div/div[2]/div/div[2]/input', train_code),
                    ('xpath=/html/body/div[1]/div/div/div/div[1]/div[2]/div/div[3]/div/div[2]/input', from_station),
                    ('xpath=/html/body/div[1]/div/div/div/div[1]/div[2]/div/div[4]/div/div[2]/input', to_station)
                ]
                save_path = os.path.join(self.dir, r'站站查询.xlsx')
                result = await self._fill_form_and_download(page, form_script, fill_actions, save_path)
                all_results.extend(result)
        elif query_type == 'plgjcx':
            date_start = params.get('date_start')
            date_end = params.get('date_end')
            id_no_list = params.get('id_no_list', [])
            # 给日期赋值
            form_script = f'''
            document.getElementsByName("startDate")[0].removeAttribute("readonly");
            document.getElementsByName("endDate")[0].removeAttribute("readonly");
            document.getElementsByName("startDate")[0].value="{date_start}";
            document.getElementsByName("endDate")[0].value="{date_end}";
            '''
            await page.evaluate(form_script)
            # 在本地生成一个 txt，把传递的 id_no_list 都写入进去
            txt_path = os.path.join(self.dir, r'id_no_list.xlsx')
            with open(txt_path, 'w') as f:
                for id_no in id_no_list:
                    f.write(id_no + '\n')
            # 上传这个 txt
            file_input = await page.wait_for_selector('input[type=file]', timeout=self.timeout)
            await file_input.set_input_files(txt_path)
            # 点击 upload 按钮
            await page.click('#upload', timeout=self.timeout)

            # 等待上传完成，可根据实际情况调整等待时间或使用更合适的等待条件
            await page.wait_for_timeout(5000)
            fill_actions = []
            save_path = r'C:\Personal\Projects\Data\关联查询.xlsx'
            result = await self._fill_form_and_download(page, form_script, fill_actions, save_path)
            all_results.extend(result)

        return all_results

    # 关联查询
    async def glcx(self, params):
        if not self.browser or await self.browser.is_closed():
            await self.init()
        page = await self.context.new_page()
        await page.goto('https://10.3.2.201:9943/rntibp/view/complex/trackQuery.html')
        results = await self._perform_query(page, 'glcx', params)
        await page.close()
        return results

    # 站站查询
    async def zzcx(self, params):
        if not self.browser or await self.browser.is_closed():
            await self.init()
        page = await self.context.new_page()
        await page.goto('https://10.3.2.201:9943/rntibp/view/complex/trackQuery.html')
        results = await self._perform_query(page, 'zzcx', params)
        await page.close()
        return results

    # 批量轨迹查询
    async def plgjcx(self, params):
        if not self.browser or await self.browser.is_closed():
            await self.init()
        page = await self.context.new_page()
        await page.goto('https://10.3.2.201:9943/rntibp/view/complex/trackQuery.html')
        results = await self._perform_query(page, 'plgjcx', params)
        await page.close()
        return results
