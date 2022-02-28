from time import sleep
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import openpyxl
import os,shutil

# driver = webdriver.Edge('msedgedriver.exe')
driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://incites.clarivate.com/zh/#/analysis/0/subject')
driver.maximize_window()
WebDriverWait(driver,20).until(EC.presence_of_element_located(('id', 'mat-input-0')))
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
driver.find_element('link text','研究方向').click()
WebDriverWait(driver,20).until(
    EC.presence_of_element_located(
    ('xpath', "//nav/ul/li/ul/li[4][@class='cl-navigation__item--secondary']")
    )
)
driver.find_element('xpath','//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/span').click()
driver.find_element('xpath', '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[2]/fieldset/ic-date-range-filter/div/span/nav/ul/li/ul/li[4]').click()
currentMin=driver.find_element('xpath','//*[@formcontrolname="currentMin"]')
currentMin.clear()
currentMin.send_keys(2010)
currentMax=driver.find_element('xpath','//*[@formcontrolname="currentMax"]')
currentMax.clear()
currentMax.send_keys(2020)
sleep(4)
driver.find_element('xpath','//*[@id="undefined"]/ic-analysis-table/section/div/div/div[1]/div/ic-add-indicator-menu/cl-popover-modal/div/cl-popover-modal-button/button').click()
sleep(3)
# 添加指标
for i in range(2,19):
    driver.find_element('xpath',f'//div/div[2]/fieldset[1]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[1]/div[1]/label/cl-text-highlight/span
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[1]/div[2]/label/cl-text-highlight/span
    # ......
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[1]/div[18]/label/cl-text-highlight/span
for i in range(3,7):
    driver.find_element('xpath',f'//div/div[2]/fieldset[2]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[2]/div[1]/label/cl-text-highlight/span
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[2]/div[2]/label/cl-text-highlight/span
    # .......
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[2]/div[7]/label/cl-text-highlight/span
for i in range(1,7):
    driver.find_element('xpath', f'//div/div[2]/fieldset[3]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[3]/div[1]/label/cl-text-highlight/span
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[3]/div[2]/label/cl-text-highlight/span
    # .....
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[3]/div[6]/label/cl-text-highlight/span
for i in range(1,7):
    driver.find_element('xpath', f'//div/div[2]/fieldset[4]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[4]/div[1]/label/cl-text-highlight/span
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[4]/div[2]/label/cl-text-highlight/span
    # ...
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[4]/div[6]/label/cl-text-highlight/span
for i in range(1, 19):
    driver.find_element('xpath', f'//div/div[2]/fieldset[5]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[5]/div[1]/label/cl-text-highlight/span
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[5]/div[2]/label/cl-text-highlight/span
    # ...
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[5]/div[18]/label/cl-text-highlight/span
for i in range(1, 7):
    driver.find_element('xpath', f'//div/div[2]/fieldset[6]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[6]/div[1]/label/cl-text-highlight/span
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[6]/div[2]/label/cl-text-highlight/span
    # ...
    # //*[@id="popover-modal-6fygvogv8"]/div/div[2]/fieldset[6]/div[6]/label/cl-text-highlight/span
for i in range(1,2):
    driver.find_element('xpath', f'//div/div[2]/fieldset[7]/div[{i}]/label/cl-text-highlight/span').click()
    # //*[@id="popover-modal-oe2rclvpc"]/div/div[2]/fieldset[7]/div[1]/label/cl-text-highlight/span
driver.find_element('xpath','//div/div[3]/button[2]').click()


# 文献类型，选择article、review
sleep(2)
driver.find_element('xpath','//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[10]/button/span').click()
sleep(1)
driver.find_element('xpath', '//*[@id="articletype"]').send_keys('article')
sleep(5)
driver.find_element('xpath', '/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/div/ul/li').click()
driver.find_element('xpath', '//*[@id="articletype"]').send_keys('review')
sleep(5)
driver.find_element('xpath', '/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/div/ul/li').click()
driver.find_element('xpath', '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
sleep(1)
# 研究方向，ESI，所有学科
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[16]/button/span').click()
sleep(1)
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[1]/div/cl-select-filter/div/fieldset/div/span/nav/ul/li/a').click()
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[1]/div/cl-select-filter/div/fieldset/div/span/nav/ul/li/ul/li[3]/label').click()
sleep(1)
driver.find_element('xpath', '//*[@id="sbjname"]').click()
sleep(1)
for j in range(1,23):
    driver.find_element('xpath', '/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter[2]/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/div/ul/li[1]').click()
    sleep(1)
driver.find_element('xpath', '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()

sleep(3)
# 机构名设置，依次选择北师大、北大、华东师大并下载每个学校的数据
for i in ["Beijing normal university","Peking university","east china normal university"]:
    driver.find_element('xpath','//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[6]/button/span').click()
    try:
        driver.find_element('xpath', '/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[1]/cl-applied-filters/div/ul/li[3]/span/button/cl-icon/span').click()
        print('已清除'+i)
    except:
        pass
    # jg_input = WebDriverWait(driver,15).until(EC.visibility_of_element_located((By.XPATH,'')))
    sleep(3)
    jg_input = driver.find_element('xpath', '//*[@id="orgname"]')
    sleep(1)
    jg_input.send_keys(i)
    sleep(5)
    # driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
    driver.find_element_by_xpath('/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[2]/div[1]/ic-analysis-sidebar/div/aside/cl-tab-nav/div/cl-tab-pane[1]/div/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/div/div/div/ul/li[1]').click()
    sleep(1)
    driver.find_element('xpath','//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
    sleep(3)
# 下载数据

    driver.execute_script("window.scrollTo(0,100);")
    WebDriverWait(driver,15).until(EC.visibility_of_element_located((By.XPATH,'/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/section/div/div/div[1]/div/ic-download-entity-data-list/cl-popover-modal/div/cl-popover-modal-button/button'))).click()
    sleep(1)
    filename = WebDriverWait(driver,15).until(EC.visibility_of_element_located((By.XPATH,'/html/body/ui-view/app/main/ui-view/ui-view-ng-upgrade/ui-view/ic-analysis/div/div[2]/div[2]/cl-tab-nav/div/cl-tab-pane[1]/div/ic-analysis-table/section/div/div/div[1]/div/ic-download-entity-data-list/cl-popover-modal/div/cl-popover-modal-content/div/div[2]/form/fieldset/div[1]/input')))
        # filename = driver.find_element('xpath', '')
    filename.clear()
    filename.send_keys(i+'所有学科数据')
    sleep(2)
    driver.execute_script("window.scrollTo(0,120);")
    driver.find_element('xpath','//div[2]/form/fieldset/div[3]/div[2]/button').click()
    sleep(3)


