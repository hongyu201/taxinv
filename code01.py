from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time, sys, os
import win32gui, win32api, win32con


chromedriver = r"chrome\chromedriver.exe"

chromeoptions = webdriver.ChromeOptions()
chromeoptions.binary_location = r"chrome\chrome.exe"

prefs = {}
chromeoptions.add_experimental_option('prefs', prefs)
chromeoptions.add_argument('log-level=3')
chromeoptions.add_argument('lang=zh_CN.UTF-8')
# chromeoptions.add_argument('--headless')
chromeoptions.add_argument('--disable-gpu')
chromeoptions.add_argument('--ignore-certificate-errors')
chromeoptions.add_argument('--user-agent="Mozilla/5.0 (Windows NT 6.1; Win64; x64) \
    AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"')

browser = webdriver.Chrome(chromedriver, options=chromeoptions)
browser.maximize_window()
browser.implicitly_wait(15)

browser.get("https://inv-veri.chinatax.gov.cn/")
time.sleep(2)
browser.execute_script("window.scrollBy(0,250)")
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

# 验证码无刷新，说明服务不可用，程序退出
yzmimg = browser.find_element_by_id("yzm_img")
yzmsrc = yzmimg.get_attribute('src')
if yzmsrc == 'https://inv-veri.chinatax.gov.cn/images/code.png':
    print("服务不可用， 退出")
    browser.quit()
    sys.exit()

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

# 进入查验结果页
browser.switch_to.window(browser.window_handles[-1])
ifrmprint = browser.find_element_by_id('dialog-body')
browser.switch_to_frame('dialog-body')
printfp = browser.find_element_by_id('printfp')
printfp.click()


# 打开打印页面， 直接发送回车
printpage = browser.find_element_by_tag_name('html')
printpage.send_keys(Keys.ENTER)

# 填入文件名， 保存
saveDlg = win32gui.FindWindow('#32770', '另存为')
parentSaveDlg = win32gui.GetParent(saveDlg)
parentSaveDlgText = win32gui.GetWindowText(parentSaveDlg)
if parentSaveDlgText != "国家税务总局全国增值税发票查验平台 - Chromium":
    print("另存为窗口查找错误， 程序退出")
    sys.exit()

saveEdit1 = win32gui.FindWindowEx(saveDlg, None, 'Edit', None)
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
    win32gui.PostMessage(saveBtn, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)


# 刷新页面，查询下一张
printpage.send_keys(Keys.ESCAPE)
browser.switch_to_window(browser.window_handles[0])
browser.get("https://inv-veri.chinatax.gov.cn/")




