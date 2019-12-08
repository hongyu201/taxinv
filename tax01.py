from selenium import webdriver
import time, os, pdfkit


# chrome选项
# basedir = os.path.dirname(__file__)
# # downloadpath = os.path.join(basedir, 'pdf')
# chromedriver = os.path.join(basedir,"chromedriver.exe")
chromedriver = "chromedriver.exe"
chromeoptions = webdriver.ChromeOptions()
prefs = {
    'profile.default_content_setting_values': {'notifications': 2},
    # 关闭保存密码提示
    'credentials_enable_service': False,
    'profile.password_manager_enabled': False,
    # 指定下载路径
    # 'download.default_directory': downloadpath,
    # 'profile.default_content_settings.popups': False,
    # 'download.prompt_for_download': False,
    # 'download.directory_upgrade': True,
    # # 不显示图片
    # 'profile.managed_default_content_settings.images': 2,
}
chromeoptions.add_experimental_option('prefs', prefs)
# chromeoptions.add_argument('--headless')
chromeoptions.add_argument('--disable-gpu')
chromeoptions.add_argument('--ignore-certificate-errors')
chromeoptions.add_argument('--user-agent="Mozilla/5.0 (Windows NT 6.1; Win64; x64) \
    AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36"')
chromeoptions.add_argument('log-level=3')
chromeoptions.add_argument('lang=zh_CN.UTF-8')
# 启动chrome
driver = webdriver.Chrome(chromedriver, options=chromeoptions)
driver.maximize_window()
# driver.set_window_size(1280, 600)
driver.implicitly_wait(10)

driver.get(r'https://inv-veri.chinatax.gov.cn/')
time.sleep(2)
driver.execute_script("window.scrollBy(0,250)")
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

# 解析发票信息
# 01,10,044001900111,73998589,33.27,20191102,80151647620902022090,0C1E,
# 01,10,044001900111,66602326,64.09,20191103,84093470454080768180,3703,
# 01,10,044001900111,64206419,23.68,20191103,83721504953813076582,7B61,
qrlst = fpqr.split(',')
tkcode = qrlst[2]
tknum = qrlst[3]
tkdate = qrlst[5]
tkamount = qrlst[4]
tkcheck = qrlst[6]
tkcheckv = tkcheck[-6:]

# 填充发票信息
fpdm = driver.find_element_by_id('fpdm') 
fphm = driver.find_element_by_id('fphm') 
kjje = driver.find_element_by_id('kjje') 
kprq = driver.find_element_by_id('kprq') 
yzm = driver.find_element_by_id('yzm') 
# 发票代码
fpdm.clear()
fpdm.send_keys(tkcode)
# 发票号码
fphm.clear()
fphm.send_keys(tknum)
# 开票日期
# kprq.clear()
driver.execute_script("arguments[0].value = " + tkdate, kprq)
# 校验码
kjje.click()
kjje.clear()
kjje.send_keys(tkcheckv)
# 验证码
yzm.clear()
yzm.click()
# 验证码手动输入
js4 = "alert('请输入正确的验证码，再点击查验按钮')"
driver.execute_script(js4)
alert = driver.switch_to.alert
while alert:
    time.sleep(1)
    try:
        alert = driver.switch_to.alert
    except:
        alert = None

# 点击打印发票
# driver.switch_to.frame("dialog-body")
# printpage = driver.page_source
# printfp = driver.find_element_by_id("printfp")
# printfp.click()
# closebt = driver.find_element_by_id("closebt")


# driver.switch_to.default_content()
# wins = driver.window_handles
# driver.switch_to.window(wins[-1])
# destinationSelect = driver.find_element_by_id("destinationSelect")


# chrome 78 关闭按钮无效，刷新页面
# driver.refresh()


# checkimg = driver.find_element_by_id("yzm_img")
# checkimg.click()

# pdfoptions
# pdfoptions = {
#     'page-size': 'A4',
#     'margin-top': '0.75in',
#     'margin-right': '0.75in',
#     'margin-bottom': '0.75in',
#     'margin-left': '0.75in',
#     'encoding': "UTF-8",
#     'custom-header' : [
#         ('Accept-Encoding', 'gzip')
#     ]
#     'cookie': [
#         ('cookie-name1', 'cookie-value1'),
#         ('cookie-name2', 'cookie-value2'),
#     ],
#     'no-outline': None
# }


