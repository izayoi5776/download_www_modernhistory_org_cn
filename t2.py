from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import TimeoutException, ElementNotVisibleException, ElementNotSelectableException
from selenium.webdriver.common.keys import Keys
import time
import os
import glob
import zipfile
from datetime import datetime
import traceback


def chkLogin(driver):
  try:
    elems = WebDriverWait(driver, 10).until(lambda x: x.find_elements_by_css_selector("input[placeholder='请输入']"))
    elems[0].send_keys("用户名")
    elems[1].send_keys("密码")
    elems[1].send_keys(Keys.RETURN)
    print("自动填写用户名密码")
  except TimeoutException:
    pass

def download(driver):
  elem = driver.find_element_by_xpath("//button/span[text()='下载']")
  actions = ActionChains(driver)
  actions.move_to_element(elem)
  actions.click(elem)
  actions.perform()
  chkLogin(driver)
  print("正在下载，等待60s")
  time.sleep(60)

def clickNextAvailableDay(driver, curdt):
  '''
  从一个月的日历中找到下一个日期。
  返回值：True移动，False没有移动（超过年上限）
  '''
  extractZips()

  # 点开日历
  elem = driver.find_element_by_css_selector("input[data-time]")
  elem.click()
  time.sleep(1.0)

  # 查找所有可用日期
  #$$("td.available div span")
  elems = driver.find_elements_by_css_selector("td.available div span")
  ret = False # 移动到下一页时候为True

  # 移动到当月的下一天
  for i in elems:
    if curdt == None or int(i.text) > curdt:
      print("选择 " + i.text + " 日")
      i.click()
      time.sleep(1)
      return True

  # $$("span.el-date-picker__header-label")[0]
  elems = driver.find_elements_by_css_selector("span.el-date-picker__header-label")
  pickerYear = elems[0].text[0:4]
  print("日历上 " + elems[0].text + elems[1].text)

  # 1950年以后就没有抗日报纸了
  # 如果当月没有可供移动的下一天，则转到下月
  if pickerYear < "1950":
    #$$("button[aria-label='下个月']")
    print("移动到下个月")
    elem = driver.find_element_by_css_selector("button[aria-label='下个月']")
    elem.click()
    time.sleep(1)
    ret = clickNextAvailableDay(driver, None)

  return ret

# ref https://stackoverflow.com/questions/35851281/python-finding-the-users-downloads-folder
def get_download_path():
    """Returns the default downloads path for linux or windows"""
    if os.name == 'nt':
        import winreg
        sub_key = r'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders'
        downloads_guid = '{374DE290-123F-4565-9164-39C4925E467B}'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub_key) as key:
            location = winreg.QueryValueEx(key, downloads_guid)[0]
        return location
    else:
        return os.path.join(os.path.expanduser('~'), 'downloads')


def extractZips():
  '''展开下载完成的zip'''
  fld = get_download_path()
  for i in glob.glob(os.path.join(fld, "泰晤士报*.zip")):
    with zipfile.ZipFile(i, 'r') as fz:
      print("展开 " + i)
      for file in fz.namelist():
        print("    " + file)
        fz.extract(file, "done")
    # 处理完的zip改名，文件名中加入时间避免重复
    newname = "done-" + datetime.now().strftime("%Y%m%d%H%M%S") + "-" + os.path.basename(i)
    print("     rename to " + newname)
    os.rename(i, os.path.join(os.path.dirname(i), newname))

def chkAlreadyDownload(fld):
  '''检查是否已经下载完毕'''
  ret = False
  chkPth = os.path.join("done", fld.replace("-", ""))
  if(os.path.exists(chkPth)):
    print(fld + " 已经存在，不需要下载")
    ret = True
  else:
    print(fld + " 未曾下载")
  return ret

def getDispDate(driver):
  '''画面上表示的时间取得'''
  #$$("input[data-time]")[0].dataset["time"]
  elem = driver.find_element_by_css_selector("input[data-time]")
  tm = elem.get_attribute("data-time")
  print("画面时间 " + tm)
  return tm

def main():
  driver = webdriver.Chrome()
  driver.implicitly_wait(time_to_wait=10)  # 只需要一个等待超时时间参数

  #driver.get('http://www.modernhistory.org.cn/reader.htm?fileCode=05563affd3cd49dbbe13a0eca62509bb&fileType=bz&startPage=1&bzId=2746994')
  #driver.get('http://www.modernhistory.org.cn/reader.htm?fileCode=05563affd3cd49dbbe13a0eca62509bb&fileType=bz&bzId=2747042&startPage=1')
  #driver.get('http://www.modernhistory.org.cn/reader.htm?fileCode=05563affd3cd49dbbe13a0eca62509bb&fileType=bz&bzId=2747449&startPage=1') # 1930-01-01
  #driver.get('http://www.modernhistory.org.cn/reader.htm?fileCode=05563affd3cd49dbbe13a0eca62509bb&fileType=bz&bzId=2747092&startPage=1') # 1930-01-29
  #driver.get('http://www.modernhistory.org.cn/reader.htm?fileCode=05563affd3cd49dbbe13a0eca62509bb&fileType=bz&bzId=2747046&startPage=1') # 1926-05-13
  driver.get('http://www.modernhistory.org.cn/reader.htm?fileCode=05563affd3cd49dbbe13a0eca62509bb&fileType=bz&bzId=2747021&startPage=1') # 1926-05-25

  tm = getDispDate(driver)
  curdt = int(tm[-2:]) # 下两衍是日期

  extractZips()
  if not chkAlreadyDownload(tm):
    download(driver)

  while clickNextAvailableDay(driver, curdt):
    time.sleep(1)
    tm = getDispDate(driver)
    curdt = int(tm[-2:]) # 下两衍是日期
    extractZips()
    time.sleep(1)
    if not chkAlreadyDownload(tm):
      download(driver)
  # driver.quit()

#====================
flg = True
while flg:
  try:
    main()
    flg = False
    break
  except:
    traceback.print_exc()
