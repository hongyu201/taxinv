from selenium import webdriver
import time


# chrome选项
chromedriver = r"chromedriver.exe"
chromeoptions = webdriver.ChromeOptions()

prefs = {
    'profile.default_content_setting_values': {'notifications': 2},
    'credentials_enable_service': False,
    'profile.password_manager_enabled': False,
    # 'profile.managed_default_content_settings.images': 2,
}
chromeoptions.add_experimental_option('prefs', prefs)
# chromeoptions.add_argument('--headless')
chromeoptions.add_argument('--disable-gpu')
chromeoptions.add_argument('--ignore-certificate-errors')
chromeoptions.add_argument('--user-agent="Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"')
chromeoptions.add_argument('log-level=3')
chromeoptions.add_argument('lang=zh_CN.UTF-8')

# 启动chrome
driver = webdriver.Chrome(chromedriver, options=chromeoptions)
driver.maximize_window()
driver.implicitly_wait(15)

driver.get(r'https://inv-veri.chinatax.gov.cn/')
time.sleep(3)

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
driver.execute_script(js1)
driver.execute_script(js2)
alert = driver.switch_to.alert
while alert:
    time.sleep(1)
    try:
        alert = driver.switch_to.alert
    except:
        alert = None

fpqr = driver.execute_script(js3)

if not fpqr:
    driver.quit()
    exit()

#解析发票信息
qrlst = fpqr.split(',')
tkcode = qrlst[2]
tknum = qrlst[3]
tkdate = qrlst[5]
tkamount = qrlst[4]
tkcheck = qrlst[6]
tkcheckv = tkcheck[-6:]

# 填充发票信息
fpdm = driver.find_element_by_id('fpdm') # 发票代码
fpdm.clear()
fpdm.send_keys(tkcode)
fphm = driver.find_element_by_id('fphm') # 发票号码
fphm.clear()
fphm.send_keys(tknum)
kprq = driver.find_element_by_id('kprq') # 开票日期
kprq.clear()
driver.execute_script("arguments[0].value = " + tkdate, kprq)
kjje = driver.find_element_by_id('kjje') # 校验码
kjje.clear()
kjje.send_keys(tkcheckv)
yzm = driver.find_element_by_id('yzm') # 验证码
yzm.clear()
yzm.click()

# chrome 78 关闭按钮无效，刷新页面
driver.refresh()


# 点击查验
# driver.switch_to.frame("dialog-body")
# checkfp = driver.find_element_by_id("checkfp")
# checkfp.click()

# 验证码失效处理
# try:
#     checkfailbtn = driver.find_element_by_id("popup_ok")
#     checkfailbtn.click()
# except:
#     pass

# checkimg = driver.find_element_by_id("yzm_img")
# checkimg.click()


# # 验证通过， 进入打印查验结果界面
# printfp = driver.find_element_by_id("printfp")
# printfp.click()

# # 关闭打印查验界面
# closebt = driver.find_element_by_id("closebt")
# closebt.click()

# 打印处理
# '''检测到查验结果页面出现后，弹出提示，自动打印'''
# wins = driver.window_handles
# if len(wins) == 2:
#     driver.switch_to.window(wins[-1])
# else:
#     pass


# cbtn = driver.find_element_by_class_name('destination-settings-change-button')
# cbtn.click()

# savepdf = driver.find_element_by_xpath("//span[text()='另存为 PDF']")
# savepdf.click()

# checkheader = driver.find_element_by_class_name('header-footer-checkbox')
# if not checkheader.is_selected():
#     checkheader.click()

# checkbg = driver.find_element_by_class_name('css-background-checkbox')
# if not checkbg.is_selected():
#     checkheader.click()

# savebtn = driver.find_element_by_xpath("//button[@class='print default']")
# savebtn.click()

