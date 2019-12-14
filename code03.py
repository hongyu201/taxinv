from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import time, sys, os
import win32gui, win32api, win32con


chromedriver = r"chrome\chromedriver.exe"
chromeoptions = webdriver.ChromeOptions()
chromeoptions.binary_location = r"chrome\chrome.exe"

prefs = {
    'profile.default_content_setting_values': {'notifications': 2},
    # 关闭保存密码提示
    'credentials_enable_service': False,
    'profile.password_manager_enabled': False,
    # 关闭未正确关闭提示
    'exit_type': 'Normal',
    'exited_cleanly': True
}
chromeoptions.add_experimental_option('prefs', prefs)
chromeoptions.add_argument('log-level=3')
chromeoptions.add_argument('lang=zh_CN.UTF-8')
# chromeoptions.add_argument('--headless')
chromeoptions.add_argument('--disable-gpu')
chromeoptions.add_argument('--ignore-certificate-errors')
chromeoptions.add_argument('--user-agent="Mozilla/5.0 (Windows NT 6.1; Win64; x64) \
    AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"')
# chromeoptions.add_argument('--disable-desktop-notifications') # 禁用桌面通知，在 Windows 中桌面通知默认是启用的
chromeoptions.add_argument('--kiosk') # 全屏模式
chromeoptions.add_argument('--no-first-run') # 跳过 Chromium 首次运行检查
# chromeoptions.add_argument('--user-data-dir=UserDataDir') # 自订使用者帐户资料夹
chromeoptions.add_argument('--user-data-dir=' + r"D:\src\chrome\chrome\userdata")
# 加载浏览器
browser = webdriver.Chrome(chromedriver, options=chromeoptions)
browser.maximize_window()
browser.implicitly_wait(15)
# 打开网页
browser.get("https://inv-veri.chinatax.gov.cn/")
time.sleep(1)
# browser.execute_script("window.scrollBy(0,250)")

# 此处插入js代码，在窗口中弹出提示信息
js1 = '''tip = document.getElementById("ktsm_tip"); \
    fpqr = document.getElementById("fpqr");
    if (!fpqr) {
        fpqr = document.createElement("p"); \
        fpqr.setAttribute("id", "fpqr"); \
        fpqr.innerHTML = ""; \
        tip.appendChild(fpqr);
    } else {
        fpqr.innerHTML = "";
    }
'''
js2 = '''p1 = document.getElementById("fpqr"); \
    p1.innerHTML = ""; \
    var fpqr = prompt('请扫描发票二维码获取发票信息'); \
    p1.innerHTML = fpqr;
'''
js3 = "return document.getElementById('fpqr').innerHTML"
fpqr = ''
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
if not fpqr:
    print("发票信息获取失败，程序退出")
    browser.close()
    browser.quit()

# 填充发票信息
qrlst = fpqr.split(',')
fpdm = browser.find_element_by_id('fpdm') # 发票代码
fpdm.send_keys(qrlst[2])
fphm = browser.find_element_by_id('fphm') # 发票号码
fphm.send_keys(qrlst[3])
kprq = browser.find_element_by_id('kprq') # 开票日期
browser.execute_script("arguments[0].value = " + qrlst[5], kprq)
kjje = browser.find_element_by_id('kjje') # 校验码
kjje.send_keys(qrlst[6][-6:])
yzm = browser.find_element_by_id('yzm') # 验证码
yzm.click()

# 验证码手动输入
js4 = "alert('请输入正确的验证码，再点击查验按钮')"
browser.execute_script(js4)
alert = browser.switch_to.alert
while alert:
    time.sleep(1)
    try:
        alert = browser.switch_to.alert
    except:
        alert = None

# 验证码无刷新，说明服务不可用，程序退出
if not alert:
    yzmimg = browser.find_element_by_id("yzm_img")
    yzmsrc = yzmimg.get_attribute('src')
    if yzmsrc == 'https://inv-veri.chinatax.gov.cn/images/code.png':
        print("验证码服务不可用，请改时间再查，5秒后程序退出")
        browser.execute_script("alert('验证码服务不可用，请改时间再查，5秒后程序退出')")
        browser.switch_to.alert()
        time.sleep(5)
        browser.quit()
        sys.exit()

# 判断查验是否成功
frmprint = None
time.sleep(5)
while True: 
    try:
        frmprint = WebDriverWait(browser, 1).until(EC.presence_of_element_located((By.ID, 'dialog-body')))
        if frmprint:
            break
        else:
            try:
                popup_message = WebDriverWait(browser, 1).until(EC.presence_of_element_located((By.ID, 'popup_message')))
                # if popup_message == "验证码失效!":
                #     # print("验证码失效!")
                # elif popup_message == "验证码错误!":
                #     # print("验证码错误!")
                # elif popup_message == "超过该张发票当日查验次数(请于次日再次查验)!":
                #     print("超过该张发票当日查验次数(请于次日再次查验)!")
                popup_ok = browser.find_element_by_id('popup_ok')
                popup_ok.click()
                browser.execute_script("alert('验证码失效、错误或超出当天查验次数，请再次尝试')")
                browser.switch_to_alert()
            except:
                pass
    except:
        pass


# 进入查验结果页, 打开打印预览页
browser.switch_to.frame('dialog-body')
btnprintfp = browser.find_element_by_id('printfp')
btnprintfp.click()
time.sleep(2)
wins = browser.window_handles
if len(wins) > 1:
    browser.switch_to.window(wins[-1])

# 打印预览页直接发送回车
printpage = browser.find_element_by_tag_name('html')
printpage.send_keys(Keys.ENTER)
time.sleep(3)

# 填入文件名， 保存
saveDlg = win32gui.FindWindow('#32770', '另存为')
parentSaveDlg = win32gui.GetParent(saveDlg)
# parentSaveDlgText = win32gui.GetWindowText(parentSaveDlg)
# if parentSaveDlgText != "国家税务总局全国增值税发票查验平台 - Chromium":
#     print("另存为窗口查找错误， 程序退出")
#     sys.exit()

winL1 = win32gui.FindWindowEx(saveDlg, None, 'DUIViewWndClassName', None)
winL2 = win32gui.FindWindowEx(winL1, None, 'DirectUIHWND', None)
winL3 = win32gui.FindWindowEx(winL2, None, 'FloatNotifySink', None)
winL4 = win32gui.FindWindowEx(winL3, None, 'ComboBox', None)
winL5 = win32gui.FindWindowEx(winL4, None, 'Edit', None)
saveFilename = os.path.join(r"D:\src\chrome\pdf", qrlst[2] + '.pdf')
win32gui.SendMessage(winL5, win32con.WM_SETTEXT, None, saveFilename)

saveBtn = win32gui.FindWindowEx(saveDlg, None, 'Button', None)
saveBtnText = win32gui.GetWindowText(saveBtn)
if saveBtnText == "保存(&S)":
    saveResult = win32gui.PostMessage(saveBtn, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
else:
    sys.exit()
time.sleep(5)

# 查询完毕， 退出
browser.switch_to.window(browser.window_handles[0])
browser.close()
browser.quit()
