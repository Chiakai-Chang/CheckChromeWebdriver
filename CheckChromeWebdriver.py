#  -*- coding: utf-8 -*-
"""
Created on Wed Jul 15 18:54:45 2020
@author: Chiakai
"""

# 不要print警告(因為requests都以verify=False連線)
import sys
import warnings  
if not sys.warnoptions:
    warnings.simplefilter("ignore")
from requests import Session
from zipfile import ZipFile
import os, time, traceback, webbrowser, platform, winreg
from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.chrome.options import Options
    

def get_chrome_version_from_registry():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
        version, _ = winreg.QueryValueEx(key, "version")
        return version
    except Exception as e:
        print(f"Error: {e}")
        return None

def get_os_type():
    system = platform.system()
    architecture = platform.architecture()[0]

    if system == "Linux":
        return "linux64"
    elif system == "Darwin":
        if architecture == "64bit" and os.uname().machine == "arm64":
            return "mac-arm64"
        else:
            return "mac-x64"
    elif system == "Windows":
        if architecture == "32bit":
            return "win32"
        else:
            return "win64"
    else:
        return "Unknown OS"

def checkDriverOutdate(chromedriverPath : str) -> bool:
    # print('>>開始檢查您的chromedriver版本是否過時')
    
    # 隱藏chrome視窗
    options = Options()
    options.headless = True
    # 這句可以讓selenium以cmd執行時不會額外print log
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    
    try:
        with webdriver.Chrome(
                service = Service(chromedriverPath),
                chrome_options = options
                ) as driver:
            driver.quit()
    except:
        x = traceback.format_exc()
        if 'This version of ChromeDriver only supports' in f'{x}':
            print('>>您的ChromeDriver版本過時囉！')
            return True
        elif 'OSError' in f'{x}':
            print('>>您的ChromeDriver不是有效的 Win32 應用程式！')
            return True
        else:
            supectFolder = []
            for f in os.listdir('c:\\'):
                if 'rogram' in f:
                    supectFolder.append(f)
            
            googleTarget = ''
            find = 0
    
            for f in supectFolder:
                path = os.path.join('c:\\',f)
                for ff in os.listdir(path):
                    if 'Google' in ff:
                        for fff in os.listdir(os.path.join(path,ff)):
                            if 'Chrome' in fff:
                                googleTarget = os.path.join(path, ff)
                                chromeFolder = os.path.join(googleTarget, 'Chrome', 'Application')
                                for ffff in os.listdir(chromeFolder):
                                    if '.' in ffff:
                                        find = 1
                                        break
            if find == 1:
                print(f'>>出錯！{x}')
            else:
                print('>>出錯，發現您電腦沒有安裝Chrome(Google瀏覽器)，要安裝以後，使用對應版本號的chromedriver才能發揮作用喔！')
                print('>>Chrome下載網址：「https://www.google.com/intl/zh-TW/chrome/」')
                _ = webbrowser.open('https://www.google.com/intl/zh-TW/chrome/')
    return False


def updateWebdriver(
        driverPath : str = os.path.join(os.getcwd(), 'data'), 
        chromeDisk : str = 'c:\\',
        MaxRetry : int = 10, 
        delZip : bool = True,
        verify: bool = False,
        retry : int = 3,
        timeout : int = 60,
        ) -> bool:
    '''
    自動分析電腦內Chrome版本，然後下載對應版本之ChromeDriver到指定之位置
    Parameters
    ----------
    driverPath : str
        想要放chromedriver之資料夾位置。預設是「os.path.join(os.getcwd(), 'data')」。
    chromeDisk : str
        chrome可能存在電腦的起始硬碟號碼。預設是「c:\\」。
    MaxRetry : int, optional
        下載過程出錯時，要重試的次數。預設是10次。
    delZip : bool, optional
        下載完成後，是否要刪除壓縮檔。預設是True。
    verify : bool, optional
        是否要SSL認證連線。預設是False(比較通用)。
    retry : int, optional
        連線重試次數。預設是3次。
    timeout : int, optional
        連線超時秒數。預設是60秒。
    Returns
    -------
    bool
        return下載是否成功的布林值
    '''

    # 確定檔案夾存在
    try:
        os.mkdir(driverPath)
    except :
        pass

    # 搜尋chrome版本號
    print('-'*30)
    print('>>開始檢查當前電腦之 Chrome 版本...')
    chromeVer = ''
    find = 0
    
    try:
        chromeVer = get_chrome_version_from_registry()
        find = 1
    except:
        supectFolder = []
        for f in os.listdir(chromeDisk):
            if 'rogram' in f:
                supectFolder.append(f)
        
        googleTarget = ''
        
        for f in supectFolder:
            path = os.path.join(chromeDisk,f)
            for ff in os.listdir(path):
                if 'Google' in ff:
                    for fff in os.listdir(os.path.join(path,ff)):
                        if 'Chrome' in fff:
                            googleTarget = os.path.join(path, ff)
                            chromeFolder = os.path.join(googleTarget, 'Chrome', 'Application')
                            try:
                                for ffff in os.listdir(chromeFolder):
                                    if '.' in ffff:
                                        chromeVer = ffff
                                        find = 1
                                        break
                            except :
                                pass
                            
    if find == 1:        
        print(f'>>您電腦當前 Chrome 瀏覽器的版本號為: {chromeVer}')
        
        # 確認是否已有Chromedriver存在
        driverPath = os.path.join(driverPath, 'chromedriver.exe')
        outdate = 0
        if os.path.exists(driverPath): # 若有就是看看是否過期
            outdate = checkDriverOutdate(driverPath)
            if outdate == 0:
                print('>>目前本機已有最新可用的 ChromeDriver！')
                return True
            else:
                # 若過期就嘗試刪除舊檔
                try:
                    os.remove(driverPath)
                except :
                    pass
        else:
            print('>>未發現 ChromeDirver！')
            print('>>請稍待...')
            
        # 沒有Chromedriver存在，或者Chromedriver已過期就重新下載、更新
        # Chrome  從 115 版起 Driver 就改與之一起發布在 the Chrome for Testing (CfT) availability dashboard
        # 參考: https://googlechromelabs.github.io/chrome-for-testing/
        # JSON API endpoints: https://github.com/GoogleChromeLabs/chrome-for-testing#json-api-endpoints
        
        url = 'https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json'
        
        # 確定本機 OS 版本
        os_version = get_os_type().strip().lower()
        
        print('>>連線獲取最新 Driver...')
        s = Session()
        s.verify = verify
        
        tempRetry = retry + 0
        count = 0
        while tempRetry > 0:
            count += 1
            try:
                r = s.get(url, timeout=timeout)
                j = r.json()
            except:
                print(f'>>獲取失敗，重新嘗試(第{count}次/最多嘗試{MaxRetry}次)')
                x = traceback.format_exc()
                print(x)
                time.sleep(1)
                tempRetry -= 1
            else:
                break
        
        version = j['channels']['Stable']['version']
        print(f'>>ChromeDriver 最新版本為: {version}')
        downloads = j['channels']['Stable']['downloads']['chromedriver']
        url_driver = ''
        for cd in downloads:
            platform = cd.get('platform', '').strip().lower()
            if os_version in platform:
                url_driver = cd['url']
                print(f'>>找到 ChromeDriver 下載連結：\n{url_driver}')
                break
        download_filename = url_driver.split('/')[-1]
        
        tempRetry = retry + 0
        count = 0
        while tempRetry > 0:
            count += 1
            
            print('>>開始嘗試下載 ChromeDriver，請稍待...', end='')
            try:
                r = s.get(url_driver, verify=False, timeout=180)
                if r.status_code != 200:
                    raise ValueError(f'連線失敗(r.status_code = {r.status_code})')
            except:
                print(f'，下載失敗，重新嘗試(第{count}次/最多嘗試{MaxRetry}次)')
                x = traceback.format_exc()
                print(x)
                time.sleep(1)
                tempRetry -= 1
            else:
                break

        with open(download_filename, 'wb') as f:
            f.write(r.content)
            
        with ZipFile(download_filename, 'r') as zip_ref:
            for file_name in zip_ref.namelist():
                if 'chromedriver.exe' in file_name:
                    # 解壓縮特定檔案到指定資料夾
                    zip_ref.extract(file_name, driverPath)
                    break  # 如果找到了，就跳出迴圈
        
        if delZip:
            try:
                os.remove(download_filename)
            except:
                x = traceback.format_exc()
                print(x)
        
        print('，完成！')
        return True
        
    else:
        print('>>發現您電腦沒有安裝 Chrome(Google瀏覽器)，要安裝以後，使用對應版本號的 ChromeDriver 才能發揮作用喔！')
        print('>>Chrome下載網址：「https://www.google.com/intl/zh-TW/chrome/」')
        _ = webbrowser.open('https://www.google.com/intl/zh-TW/chrome/')
        return False
    
    return False

if __name__ == '__main__':
    result = updateWebdriver()