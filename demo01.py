from selenium import webdriver
import time, os


# chrome选项
basedir = os.path.dirname(__file__)
downloadpath = os.path.join(basedir, 'pdf')
chromedriver = r"chromedriver.exe"
chromeoptions = webdriver.ChromeOptions()
prefs = {
    'profile.default_content_setting_values': {'notifications': 2},
    'credentials_enable_service': False,
    'profile.password_manager_enabled': False,
    'download.default_directory': downloadpath,
    'profile.default_content_settings.popups': False,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'credentials_enable_service': False,
    'profile.password_manager_enable': False
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
# caps 设定默认下载路径downloadsPath
# 启动chrome
driver = webdriver.Chrome(chromedriver, options=chromeoptions)
driver.maximize_window()
# driver.set_window_size(1280, 600)
driver.implicitly_wait(15)

driver.get(r'https://inv-veri.chinatax.gov.cn/')