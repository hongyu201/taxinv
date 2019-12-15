from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time
import sys
import os
import logging
import win32gui
import win32con


# logging config
logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fileH = logging.FileHandler("log.txt")
fileH.setLevel(logging.INFO)
fileH.setFormatter(formatter)
streamH = logging.StreamHandler()
streamH.setLevel(logging.INFO)
streamH.setFormatter(formatter)
logger.addHandler(fileH)
logger.addHandler(streamH)
# chrome config
basedir = os.getcwd()
chromedriver = os.path.join(basedir, "chrome", "chromedriver.exe")
prefs = {
    'profile.default_content_setting_values': {'notifications': 2},
    # 关闭保存密码提示
    'credentials_enable_service': False,
    'profile.password_manager_enabled': False,
    # 关闭未正确关闭提示
    'exit_type': 'Normal',
    'exited_cleanly': True
}
chromeoptions = webdriver.ChromeOptions()
chromeoptions.binary_location = os.path.join(basedir, "chrome", "chrome.exe")
chromeoptions.add_experimental_option('prefs', prefs)
chromeoptions.add_argument('log-level=3')
chromeoptions.add_argument('lang=zh_CN.UTF-8')
# chromeoptions.add_argument('--headless')
chromeoptions.add_argument('--disable-gpu')
chromeoptions.add_argument('--ignore-certificate-errors')
chromeoptions.add_argument('--user-agent="Mozilla/5.0 (Windows NT 6.1; Win64; x64) \
    AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"')
# chromeoptions.add_argument('--disable-desktop-notifications') # 禁用桌面通知，在 Windows 中桌面通知默认是启用的
# chromeoptions.add_argument('--kiosk') # 全屏模式
chromeoptions.add_argument('--no-first-run')  # 跳过 Chromium 首次运行检查
# chromeoptions.add_argument('--user-data-dir=UserDataDir') # 自订使用者帐户资料夹
chromeoptions.add_argument(
    '--user-data-dir=' + os.path.join(basedir, "chrome", "userdata"))
chromeoptions.add_argument('--incognito') # 隐身模式/无痕模式
# 加载浏览器
browser = webdriver.Chrome(chromedriver, options=chromeoptions)
browser.maximize_window()
browser.implicitly_wait(10)
# 退出代码


def invExit():
    browser.close()
    browser.quit()
    sys.exit()


# 进入发票查验循环
while True:
    try:
        # logger.info(basedir)
        # logger.info(chromedriver)
        # logger.info(chromeoptions.binary_location)
        # 打开国家税务局网站
        browser.get("https://inv-veri.chinatax.gov.cn/")
        browser.execute_script("window.scrollBy(0,250)")
        logger.debug("start inv-veri")
        # 此处插入js代码，在浏览器中获得发票信息

        def getInvinfo():
            js1 = '''tip = document.getElementById("ktsm_tip"); \
                fpqr = document.getElementById("fpqr");
                if (!fpqr) {
                    fpqr = document.createElement("p"); \
                    fpqr.setAttribute("id", "fpqr"); \
                    fpqr.innerHTML = ""; \
                    tip.appendChild(fpqr);
                } else {
                    fpqr.innerHTML = "";
                }'''
            js2 = '''p1 = document.getElementById("fpqr"); \
                p1.innerHTML = ""; \
                var fpqr = prompt('请扫描发票二维码获取发票信息', ''); \
                p1.innerHTML = fpqr;'''
            js3 = "return document.getElementById('fpqr').innerHTML"
            browser.execute_script(js1)
            browser.execute_script(js2)
            alert = browser.switch_to.alert

            while alert:
                time.sleep(1)
                try:
                    alert = browser.switch_to.alert
                except:
                    alert = None

            fpqr = browser.execute_script(js3)
            return fpqr

        fpqr = getInvinfo()
        while not fpqr:
            fpqr = getInvinfo()

        qrlst = fpqr.split(',')
        # 填充发票信息

        def setInvinfo(qrlst):
            fpdm = browser.find_element_by_id('fpdm')  # 发票代码
            fpdm.clear()
            fpdm.send_keys(qrlst[2])
            fphm = browser.find_element_by_id('fphm')  # 发票号码
            fphm.clear()
            fphm.send_keys(qrlst[3])
            kprq = browser.find_element_by_id('kprq')  # 开票日期
            browser.execute_script("arguments[0].value = " + qrlst[5], kprq)
            kjje = browser.find_element_by_id('kjje')  # 校验码
            kjje.clear()
            kjje.send_keys(qrlst[6][-6:])
            yzm = browser.find_element_by_id('yzm')  # 验证码
            yzm.clear()
            yzm.click()

        setInvinfo(qrlst)
        # 手动输入验证码
        # 检查验证码服务，不可用时直接退出程序
        time.sleep(2)

        def checkService():
            yzmimg = browser.find_element_by_id("yzm_img")
            yzmsrc = yzmimg.get_attribute('src')
            if yzmsrc == 'https://inv-veri.chinatax.gov.cn/images/code.png':
                tip = r"验证码服务不可用，请改时间再查"
                logger.warning(tip)
                browser.execute_script("alert(arguments[0])", tip)
                time.sleep(5)
                browser.switch_to.alert.accept()
                invExit()

        checkService()
        # 检查发票查验是否成功
        def checkVeri():
            frmprint = None
            try:
                WebDriverWait(browser, 1).until(
                    EC.presence_of_element_located((By.ID, 'dialog-body')))
                frmprint = browser.find_element_by_id('dialog-body')
            except Exception as e:
                logger.debug(e)
                if not frmprint:
                    try:
                        WebDriverWait(browser, 5).until(
                            EC.presence_of_element_located((By.ID, 'popup_message')))
                        popup_message = browser.find_element_by_id(
                            'popup_message')
                        if popup_message:
                            logger.debug(popup_message.text)
                            browser.find_element_by_id('popup_ok').click()
                            time.sleep(2)
                            fpqr = getInvinfo()
                            qrlst = fpqr.split(',')
                            setInvinfo(qrlst)
                            checkService()
                    except Exception as e:
                        logger.debug(e)
            finally:
                return frmprint

        while True:
            frmprint = checkVeri()
            if frmprint:
                logger.info('verify succeesful, inv code:' + qrlst[2])
                break
            else:
                time.sleep(1)
                continue
        # 发票查验成功后, 打开打印预览页
        time.sleep(1)
        browser.execute_script("window.scrollBy(0,0)")
        browser.switch_to.frame('dialog-body')
        cycs = browser.find_element_by_id('cycs').text[-3:]
        invname = "{}-{}-{}-{}.pdf".format(qrlst[2], qrlst[3], qrlst[5], cycs)
        logger.debug(invname)
        curHans = browser.window_handles
        btnprintfp = browser.find_element_by_id('printfp')
        btnprintfp.click()
        WebDriverWait(browser, 5).until(EC.new_window_is_opened(curHans))
        time.sleep(3)  # 打印预览窗口延时
        wins = browser.window_handles
        if len(wins) == 2:
            browser.switch_to.window(wins[-1])
            htmlprint = browser.find_element_by_tag_name('html')
            htmlprint.click()
            htmlprint.send_keys(Keys.ENTER)
        else:
            logger.warning("无法切换到打印预览窗口，程序退出")
            invExit()

        # 填入文件名， 保存
        time.sleep(1)  # 另存为窗口延时
        saveFilename = os.path.join(basedir, "pdf", invname)
        logger.debug(saveFilename)
        saveDlg = win32gui.FindWindow('#32770', '另存为')
        winL1 = win32gui.FindWindowEx(saveDlg, None, 'DUIViewWndClassName', None)
        winL2 = win32gui.FindWindowEx(winL1, None, 'DirectUIHWND', None)
        winL3 = win32gui.FindWindowEx(winL2, None, 'FloatNotifySink', None)
        winL4 = win32gui.FindWindowEx(winL3, None, 'ComboBox', None)
        winL5 = win32gui.FindWindowEx(winL4, None, 'Edit', None)
        win32gui.SendMessage(winL5, win32con.WM_SETTEXT, None, saveFilename)
        saveBtn = win32gui.FindWindowEx(saveDlg, None, 'Button', None)
        win32gui.PostMessage(saveBtn, win32con.WM_KEYDOWN,
                             win32con.VK_RETURN, 0)
        logger.info('save successfully. file:' + saveFilename)

        # 查询完毕， 退出
        time.sleep(2)
        browser.switch_to.window(browser.window_handles[0])
        logger.info('inv {} verify done.'.format(qrlst[2]))
    except Exception as e:
        logger.warning(e)
        invExit()
