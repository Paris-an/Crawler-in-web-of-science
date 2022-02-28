from time import sleep
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import os,shutil

# driver = webdriver.Edge('msedgedriver.exe')
driver = webdriver.Chrome('chromedriver.exe')
#创建driver对象
driver.get('https://incites.clarivate.com/zh/#/analysis/0/organization')
#打开目标url
driver.maximize_window()
#窗口最大化
WebDriverWait(driver,10).until(EC.presence_of_element_located(('id', 'mat-input-0')))
#等待
driver.find_element('id', 'mat-input-0').send_keys('XXX')
#输入账号，引号内的xxx更换为自己在该网站注册的邮箱
driver.find_element('id', 'mat-input-1').send_keys('XXXXX')
#输入密码，引号内的xxxxx更换为自己在该网站注册的邮箱的密码
driver.find_element('id', 'signIn-btn').click()
#登录
try:
    WebDriverWait(driver, 100).until(
        EC.presence_of_element_located(
            ('id', 'onetrust-accept-btn-handler')
        )
    )
    driver.execute_script('arguments[0].click()',driver.find_element('id', 'onetrust-accept-btn-handler'))
except:
    pass
#选择弹出选项
sleep(3)
driver.find_element('link text', '分析').click()
#点击分析
driver.find_element('link text','机构').click()
#点击机构
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
#设置年份
sleep(4)
driver.find_element('xpath','//*[@id="undefined"]/ic-analysis-table/section/div/div/div[1]/div/ic-add-indicator-menu/cl-popover-modal/div/cl-popover-modal-button/button').click()
sleep(3)
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
#选择全部指标
driver.find_element('xpath','//div/div[3]/button[2]').click()
sleep(2)
driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[1]/button/span').click()
jg_input = driver.find_element('xpath', '//*[@id="orgname"]')
sleep(3)
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
#更新机构名称
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[9]/button/span').click()
sleep(1)
driver.find_element('xpath', '//*[@id="articletype"]').send_keys('article')
sleep(3)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
driver.find_element('xpath', '//*[@id="articletype"]').send_keys('review')
sleep(3)
driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
driver.find_element('xpath',
                    '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
sleep(1)
#更新文献类型
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
#更新研究方向
def zz(dq,mkjg):
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[1]/button/span').click()
    sleep(1)
    driver.find_element('xpath', '//*[@id="orgname"]').send_keys(mkjg)
    sleep(2)
    try:
        driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
    except:
        print(f'{dq}学科未下载{mkjg}数据，用清华大学代替，需重新下载')
        sleep(1)
        driver.find_element('xpath', '//*[@id="orgname"]').clear()
        driver.find_element('xpath', '//*[@id="orgname"]').send_keys('Tsinghua university')
        sleep(2)
        driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
    sleep(1)
    driver.find_element('xpath','//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[18]/button/span[1]').click()
    sleep(1)
    driver.find_element('xpath', '//*[@id="sbjname"]').send_keys(dq)
    sleep(2)
    # try:
    driver.find_element('xpath', '//*[@id="cl-multi-select__options__item--0"]').click()
    # except:
    # print(f'{dq}学科未下载{mkjg}数据')
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
#读表重更新
def down1(dq,row):
    driver.execute_script("window.scrollTo(0,100);")
    driver.find_element('xpath',
                        '//*[@id="undefined"]/ic-analysis-table/section/div/div/div[1]/div/ic-download-entity-data-list/cl-popover-modal/div/cl-popover-modal-button/button').click()
    sleep(1)
    filename = driver.find_element('xpath', '//*[@id="fileName"]')
    filename.clear()
    name1=row+'-'+dq
    filename.send_keys(name1)
    sleep(2)
    driver.execute_script("window.scrollTo(0,120);")
    driver.find_element('xpath','//div[2]/form/fieldset/div[3]/div[2]/button').click()
    sleep(10)
    try:
        shutil.move(path + '\\' + name1 + '.csv', newpath1)
    except:
        f_names=os.listdir(path)
        for f_name in f_names:
        # print(f_name)
            f_name_1=f_name.replace('.csv','')
            try:
                shutil.move(path + '\\' + f_name_1 +'.csv', newpath1)
                print("【"+f_name_1,".csv】文件已移动到'Incites指标数据'文件夹内")
            except:
                pass
    #移动到对应文件夹
#下载指标数据
def down2(dq,mkjg,row):
    for i in range(1,5):
        name = driver.find_element('xpath',
                                   f'//table/tbody[3]/tr[{i}]/td[1]/cl-sortable-table-cell/span/span[2]/cl-text-highlight/span').text
        name2=row+'-'+name
        driver.find_element('xpath', f'//table/tbody[3]/tr[{i}]/td[2]/cl-sortable-table-cell/span/span/button').click()
        sleep(3)
        driver.find_element('xpath','//*[@id="undefined"]/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[2]/div[2]/ic-document-list/div[1]/span[4]/ic-document-list-download/cl-popover-modal/div/cl-popover-modal-button/button').click()
        name_input=driver.find_element('xpath','//*[@id="file-name"]')
        name_input.clear()
        name_input.send_keys(name2)
        sleep(1)
        driver.find_element('xpath','//div[2]/form/fieldset/div[3]/button').click()
        sleep(1)
        driver.find_element('xpath',
                            '//*[@id="undefined"]/ic-analysis-table/ic-row-details-modal/cl-overlay-modal/cl-overlay-modal-content/div/div/div[1]/button').click()
        sleep(10)
        shutil.move(path + '\\' + name2 + '.csv', newpath2)
        #移动到对应文件夹
    sleep(4)
#下载论文数据
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[1]/button/span').click()
    sleep(2)
    try:
        driver.find_element('xpath',f'// *[ @ id = "analysis-sidebar-filters"]//span[translate(@title, "abcdefghijklmnopqrstuvwxyz","ABCDEFGHIJKLMNOPQRSTUVWXYZ")="{mkjg}"]/../button').click()
        sleep(2)
    except:
        driver.find_element('xpath',f'//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[3]/cl-analysis-filter/div/cl-search-filter/div/cl-multi-select/div/div[2]/ul/li[4]/button').click()
        sleep(2)
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
    sleep(1)
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[1]/div[3]/ul/li[18]/button/span[1]').click()
    sleep(2)
    try:
        driver.find_element('xpath',
                        f'// *[ @ id = "analysis-sidebar-filters"]//span[translate(@title, "abcdefghijklmnopqrstuvwxyz","ABCDEFGHIJKLMNOPQRSTUVWXYZ")="{dq}"]/../button').click()
        sleep(1)
    except:
        pass
    driver.find_element('xpath',
                        '//*[@id="analysis-sidebar-filters"]/ic-sidebar-filters/div[2]/ic-sidebar-filter/div/div[4]/div[1]/button[2]').click()
#重置上一条

wb = openpyxl.load_workbook('各学科门槛机构.xlsx')
sheets = wb.active
path=r'C:\Users\Administrator\Downloads'
newpath1=path+r'\Incites指标数据'
try:
    os.makedirs(newpath1)
except:
    pass

for sheet in sheets[2:25]:
    #设置爬取的机构门槛行数范围，如因网速中断后，可从中断的行数爬，不必从头开始
    ts='Tsinghua University'
    dq=sheet[0].value
    mkjg=sheet[2].value
    row = str(sheet[2].row)
    newpath2 = path + r'\各学科我校及对标机构论文详细信息' + '\\' + row+'-'+dq
    try:
        os.makedirs(newpath2)
    except:
        pass
    zz(dq,mkjg)
    sleep(2)
    down1(dq,row)
    sleep(2)
    down2(dq,mkjg,row)
    sleep(1)

