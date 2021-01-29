# -*- coding: utf-8 -*-
"""
Created on Wed Jul 15 18:54:45 2020

@author: Chiakai
"""
#不要print警告(因為requests都以verify=False連線)
import sys
import warnings  
if not sys.warnoptions:
    warnings.simplefilter("ignore")
from requests_html import HTMLSession
from zipfile import ZipFile
import os, time, traceback, webbrowser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
    

def checkDriverOutdate(chromedriverPath : str) -> bool:
    #print('>>開始檢查您的chromedriver版本是否過時')
    
    #隱藏chrome視窗
    options = Options()
    options.headless = True
    #這句可以讓selenium以cmd執行時不會額外print log
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    
    try:
        with webdriver.Chrome(
                executable_path=chromedriverPath,
                chrome_options=options
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

    #確定檔案夾存在
    try:
        os.mkdir(driverPath)
    except :
        pass

    #搜尋chrome版本號
    print('>>搜尋本機目前Chrome版本號')
    supectFolder = []
    for f in os.listdir(chromeDisk):
        if 'rogram' in f:
            supectFolder.append(f)
    
    googleTarget = ''
    chromeVer = ''
    find = 0
    
    for f in supectFolder:
        path = os.path.join(chromeDisk,f)
        for ff in os.listdir(path):
            if 'Google' in ff:
                for fff in os.listdir(os.path.join(path,ff)):
                    if 'Chrome' in fff:
                        googleTarget = os.path.join(path, ff)
                        chromeFolder = os.path.join(googleTarget, 'Chrome', 'Application')
                        for ffff in os.listdir(chromeFolder):
                            if '.' in ffff:
                                chromeVer = ffff
                                find = 1
                                break
                            
    print(f'>>您電腦當前Chrome瀏覽器的版本號為：{chromeVer}')
    
    #確認是否已有Chromedriver存在
    driverPath = os.path.join(driverPath, 'chromedriver.exe')
    outdate = 0
    if os.path.exists(driverPath): #若有就是看看是否過期
        outdate = checkDriverOutdate(driverPath)
        if outdate == 0:
            print('>>目前本機已有最新可用的ChromeDriver！')
            return True
        else:
            #若過期就嘗試刪除舊檔
            try:
                os.remove(driverPath)
            except :
                pass
        
    #沒有Chromedriver存在，或者Chromedriver已過期就重新下載、更新
    if find == 1:
        print(f'>>找到您目前使用的Chrome版本號為：{chromeVer}')
        myVerList = chromeVer.split('.')
        #下載符合版本號之driver
        s = HTMLSession()
        s.verify = verify
        
        #先抓版本號最近的
        url = 'https://chromedriver.chromium.org/downloads'
        
        tempRetry = retry + 0
        while tempRetry > 0:
            try:
                r = s.get(url, timeout=timeout)
                
                links = r.html.links
                
                matchList = []
                closeList = []
                
                for link in links:
                    if '?path=' in link:
                        count = 0
                        for string in myVerList:
                            if string in link:
                                count += 1
                        if count == 4:
                            #生成確實的下載網址再append
                            link_fix = str(link).replace('index.html?path=','') + 'chromedriver_win32.zip'
                            matchList.append(link_fix)
                        elif count > 2:
                            #生成確實的下載網址再append
                            link_fix = str(link).replace('index.html?path=','') + 'chromedriver_win32.zip'
                            closeList.append(link_fix)
            except :
                tempRetry -= 1
                x = traceback.format_exc()
                print(x)
                print('>>連線出錯！問題如上...')
                if tempRetry > 0:
                    print('>>重新再嘗試一次...')
                else:
                    raise ConnectionError(f'目前無法連線「{url}」，或連線出錯！')
        
        if len(matchList) > 0:
            url_driver = matchList[0]
            tempRetry = retry + 0
            count = 0
            while tempRetry > 0:
                count += 1
                print(f'找到driver下載連結如下：\n{url_driver}')
                print('>>開始嘗試下載chromedriver', end='')
                try:
                    r = s.get(url_driver, verify=False, timeout=180)
                    break
                except:
                    print(f'，下載失敗，重新嘗試(第{count}次/最多嘗試{MaxRetry}次)')
                    x = traceback.format_exc()
                    print(x)
                    time.sleep(1)
                    tempRetry -= 1
            if r.status_code == 200:
                pass
            else:
                for i in range(len(closeList)-1, -1, -1):
                    url_driver = closeList[i]
                    retry = MaxRetry + 0
                    count = 0
                    while retry > 0:
                        count += 1
                        print(f'找到driver下載連結如下：\n{url_driver}')
                        print('>>開始嘗試下載ChromeDriver', end='')
                        try:
                            r = s.get(url_driver, verify=False, timeout=180)
                            break
                        except:
                            print(f'，下載失敗，重新嘗試(第{count}次/最多嘗試{MaxRetry}次)')
                            x = traceback.format_exc()
                            print(x)
                            time.sleep(1)
                            retry -= 1
                    if r.status_code == 200:
                        break
        else:
            for i in range(len(closeList)-1, -1, -1):
                url_driver = closeList[i]
                tempRetry = retry + 0
                count = 0
                while tempRetry > 0:
                    count += 1
                    print(f'找到driver下載連結如下：\n{url_driver}')
                    print('>>開始嘗試下載ChromeDriver', end='')
                    try:
                        r = s.get(url_driver, verify=False, timeout=180)
                        break
                    except:
                        print(f'，下載失敗，重新嘗試(第{count}次/最多嘗試{MaxRetry}次)')
                        x = traceback.format_exc()
                        print(x)
                        time.sleep(1)
                        tempRetry -= 1
                if r.status_code == 200:
                    break
                
        if r.status_code == 200:
            with open('chromedriver_win32.zip', 'wb') as f:
                f.write(r.content)
            
            with ZipFile('chromedriver_win32.zip', mode='r') as z:
                if 'chromedriver.exe' in driverPath:
                    driverPath = os.path.dirname(driverPath )
                z.extractall(driverPath)
            
            if delZip:
                os.remove('chromedriver_win32.zip')
            
            print('，完成！')
            return True
        else:
            print('>>發現下載最新版ChromeDriver失敗！可能找不到適合您電腦的版本')
            return False
        
    else:
        print('>>發現您電腦沒有安裝Chrome(Google瀏覽器)，要安裝以後，使用對應版本號的chromedriver才能發揮作用喔！')
        print('>>Chrome下載網址：「https://www.google.com/intl/zh-TW/chrome/」')
        _ = webbrowser.open('https://www.google.com/intl/zh-TW/chrome/')
        return False
    
    return False

if __name__ == '__main__':
    result = updateWebdriver()