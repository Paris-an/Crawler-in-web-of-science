from time import sleep
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import os,shutil

# driver = webdriver.Edge('msedgedriver.exe')
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://incites.clarivate.com/zh/#/analysis/0/organization')
driver.maximize_window()
WebDriverWait(driver,10).until(EC.presence_of_element_located(('id', 'mat-input-0')))
# 引号内的xxx更换为自己在该网站注册的邮箱
driver.find_element('id', 'mat-input-0').send_keys('XXX')
# 引号内的xxxxx更换为自己在该网站注册的邮箱的密码
driver.find_element('id', 'mat-input-1').send_keys('XXXXX')
driver.find_element('id', 'signIn-btn').click()
try:
    WebDriverWait(driver, 100).until(
        EC.presence_of_element_located(
            ('id', 'onetrust-accept-btn-handler')
        )
    )
    driver.execute_script('arguments[0].click()',driver.find_element('id', 'onetrust-accept-btn-handler'))
except:
    pass
sleep(3)
driver.find_element('link text', '分析').click()
driver.find_element('link text','机构').click()
WebDriverWait(driver,20).until(
    EC.presence_of_element_located(
    ('xpath', "//nav/ul/li/ul/li[4][@class='cl-navigation__item--secondary']")
    )
)
driver.find_element('xpath','//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/span').click()
driver.find_element('xpath', '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/span/nav/ul/li/ul/li[4]').click()
currentMin=driver.find_element('xpath','//*[@formcontrolname="currentMin"]')
currentMin.clear()
currentMin.send_keys(2011)
currentMax=driver.find_element('xpath','//*[@formcontrolname="currentMax"]')
currentMax.clear()
currentMax.send_keys(2021)
sleep(4)
driver.find_element('xpath','//*[@id="undefined"]/ic-analysis-table/section/div/div/div[1]/div/ic-add-indicator-menu/cl-popover-modal/div/cl-popover-modal-button/button').click()
sleep(3)
# 添加指标
for i in range(2,20):
    driver.find_element('xpath',f'//div/div[2]/fieldset[1]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(4,11):
    driver.find_element('xpath',f'//div/div[2]/fieldset[2]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(1, 9):
    driver.find_element('xpath', f'//div/div[2]/fieldset[3]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(1, 14):
    driver.find_element('xpath', f'//div/div[2]/fieldset[4]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(1, 7):
    driver.find_element('xpath', f'//div/div[2]/fieldset[5]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(1, 19):
    driver.find_element('xpath', f'//div/div[2]/fieldset[6]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(1, 7):
    driver.find_element('xpath', f'//div/div[2]/fieldset[7]/div[{i}]/label/cl-text-highlight/span').click()
for i in range(1, 7):
    if i != 3:
        driver.find_element('xpath', f'//div/div[2]/fieldset[8]/div[{i}]/label/cl-text-highlight/span').click()
sleep(2)
driver.find_element('xpath','//div/div[3]/button[2]').click()
sleep(2)
driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[1]/button/span').click()
# 机构名设置，选择北师大、北大、华东师大
jg_input = driver.find_element('xpath', '//*[@id="orgname"]')
jg_input.send_keys('Beijing normal university')
sleep(3)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
sleep(1)
jg_input.send_keys('Peking university')
sleep(3)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
sleep(1)
jg_input.send_keys('east china normal university')
sleep(3)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
sleep(1)
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
sleep(1)
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[9]/button/span').click()
sleep(1)
# 文献类型，选择article、review
driver.find_element('xpath', '//*[@id="articletype"]').send_keys('article')
sleep(5)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
driver.find_element('xpath', '//*[@id="articletype"]').send_keys('review')
sleep(5)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
sleep(1)
#研究方向
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[18]/button/span').click()
sleep(1)
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[1]/div/cl-select-filter/div/fieldset/div/span/nav/ul/li/a').click()
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[1]/div/cl-select-filter/div/fieldset/div/span/nav/ul/li/ul/li[3]/label').click()
sleep(1)

driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()

sleep(1)
def down1():
    driver.execute_script("window.scrollTo(0,100);")
    driver.find_element('xpath',
                        '//*[@id="undefined"]/ic-analysis-table/section/div/div/div[1]/div/ic-download-entity-data-list/cl-popover-modal/div/cl-popover-modal-button/button').click()
    sleep(1)
    filename = driver.find_element('xpath', '//*[@id="fileName"]')
    filename.clear()
    filename.send_keys('北师大及机构整体对标高校incites数据')
    sleep(2)
    driver.execute_script("window.scrollTo(0,120);")
    driver.find_element('xpath','//div[2]/form/fieldset/div[3]/div[2]/button').click()
    sleep(3)
path=r'C:\Users\Administrator\Downloads'
try:
    os.makedirs(path)
except:
    pass
down1()

