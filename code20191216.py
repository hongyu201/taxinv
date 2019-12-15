from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
# from selenium.webdriver.support.events import EventFiringWebDriver, AbstractEventListener
import time
import sys
import os
import signal
import logging
import win32gui
import win32con


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
# chromeoptions.add_argument('--disable-desktop-notifications') # 禁用桌面通知
# chromeoptions.add_argument('--kiosk') # 全屏模式
chromeoptions.add_argument('--no-first-run')  # 跳过 Chromium 首次运行检查
chromeoptions.add_argument('--user-data-dir=' + os.path.join(basedir, "chrome", "userdata"))
chromeoptions.add_argument('--incognito') # 隐身模式/无痕模式
# logging config
logger = logging.getLogger()
logger.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s-%(name)s-%(levelname)s-%(message)s')
fileH = logging.FileHandler("log.txt")
fileH.setLevel(logging.INFO)
fileH.setFormatter(formatter)
streamH = logging.StreamHandler()
streamH.setLevel(logging.INFO)
streamH.setFormatter(formatter)
logger.addHandler(fileH)
logger.addHandler(streamH)

# 进入发票查验循环
try:
    # 加载浏览器
    browser = webdriver.Chrome(chromedriver, options=chromeoptions)
    browser.maximize_window()
    browser.implicitly_wait(10)
    # 退出代码
    def invExit(browser):
        browser.close()
        browser.quit()
        sys.exit()

    while True:
        # 打开国家税务局网站
        browser.get("https://inv-veri.chinatax.gov.cn/")
        time.sleep(2)
        browser.execute_script("window.scrollBy(0,250)")
        # 获得发票信息
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
        qrlst = fpqr.split(',')
        # 此处缺少发票二维码信息的校验
        # 填充发票信息
        def setInvinfo(browser, qrlst):
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
            time.sleep(1)
        
        setInvinfo(browser, qrlst)
        # 检查验证码服务，不可用时直接退出程序
        def checkService(browser):
            yzmimg = browser.find_element_by_id("yzm_img")
            yzmsrc = yzmimg.get_attribute('src')
            if yzmsrc == 'https://inv-veri.chinatax.gov.cn/images/code.png':
                return False
            else:
                return True

        if not checkService(browser):
            tip = r"验证码服务不可用，请改时间再查, 5秒后程序退出"
            logger.warning(tip)
            browser.execute_script("alert(arguments[0])", tip)
            time.sleep(5)
            browser.switch_to.alert.accept()
            invExit(browser)
        # 手动输入验证码
        time.sleep(2)
        # 检查发票查验是否成功
        frmprint = None
        popup_message = None
        while not frmprint:
            try:
                WebDriverWait(browser, 1).until(EC.presence_of_element_located((By.ID, 'dialog-body')))
                frmprint = browser.find_element_by_id('dialog-body')
            except Exception as e:
                # logger.debug(e)
                pass

            if not frmprint:
                try:
                    WebDriverWait(browser, 1).until(EC.presence_of_element_located((By.ID, 'popup_message')))
                    popup_message = browser.find_element_by_id('popup_message')
                    if popup_message.text == "验证码失效!":
                        continue
                    elif popup_message.text == "验证码错误!":
                        continue
                    # elif popup_message.text == "超过该张发票当日查验次数(请于次日再次查验)!":
                    #     continue
                    else:
                        browser.execute_script("alert('请换发票查询, 5秒后页面自动刷新')")
                        time.sleep(5)
                        browser.switch_to.alert.accept()
                        browser.refresh()
                        time.sleep(2)
                        browser.execute_script("window.scrollBy(0,250)")
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
                        qrlst = fpqr.split(',')
                        setInvinfo(browser, qrlst)
                except Exception as e:
                    # logger.debug(e)
                    continue

        # 发票查验成功后, 打印查验结果
        def invPrint(browser):
            time.sleep(1)
            invName = ''
            browser.execute_script("window.scrollBy(0,0)")
            browser.switch_to.frame('dialog-body')
            cycs = browser.find_element_by_id('cycs').text[-3:]
            invName = "{}-{}-{}-{}.pdf".format(qrlst[2], qrlst[3], qrlst[5], cycs)
            logger.debug(invName)
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
                invExit(browser)
            
            return invName
        
        invName = invPrint(browser)
        # 填入文件名， 保存查询结果文件
        def invSave(invName):
            time.sleep(1)  # 另存为窗口延时
            saveFilename = os.path.join(basedir, "pdf", invName)
            logger.debug(saveFilename)
            saveDlg = win32gui.FindWindow('#32770', '另存为')
            winL1 = win32gui.FindWindowEx(saveDlg, None, 'DUIViewWndClassName', None)
            winL2 = win32gui.FindWindowEx(winL1, None, 'DirectUIHWND', None)
            winL3 = win32gui.FindWindowEx(winL2, None, 'FloatNotifySink', None)
            winL4 = win32gui.FindWindowEx(winL3, None, 'ComboBox', None)
            winL5 = win32gui.FindWindowEx(winL4, None, 'Edit', None)
            win32gui.SendMessage(winL5, win32con.WM_SETTEXT, None, saveFilename)
            saveBtn = win32gui.FindWindowEx(saveDlg, None, 'Button', None)
            win32gui.PostMessage(saveBtn, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            logger.info('save successfully. file:' + saveFilename)

        invSave(invName)
        # 本次查询完毕， 进入新的查询
        logger.info('inv {} verify done.'.format(qrlst[2]))
        time.sleep(2)
        browser.switch_to.window(browser.window_handles[0])
except Exception as e:
    logger.warning(e)
    browser.close()
    browser.quit()
    sys.exit()
