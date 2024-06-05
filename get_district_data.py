import time

import pandas as pd
from PIL import Image
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# need to paid
from twocaptcha import TwoCaptcha

# 建立在截圖中定位captcha的function
def locate_captcha(element, scroll_down, scaling_factor):
    left = int(element.location['x'] * scaling_factor)
    right = int((element.location['x'] + element.size['width']) * scaling_factor)
    top = int((element.location['y'] - scroll_down) * scaling_factor)
    bottom = int((element.location['y'] + element.size['height'] - scroll_down) * scaling_factor)
    return left, top, right, bottom

# 建立處理captcha的function
def captcha_solver(driver):
    print("定位captcha中")
    # 找出captcha
    captcha_element = driver.find_element(By.XPATH, '//*[@id="captchaImage_captchaKey"]')
    
    # 下拉到captcha的javascript程式碼
    driver.execute_script("arguments[0].scrollIntoView();", captcha_element)
    
    # 回報對應下拉距離
    scroll_down = driver.execute_script("return window.scrollY;")
    print("截取當前網頁畫面")
    # 截取當前畫面
    captcha_screenshot = driver.save_screenshot('./captcha_screenshot.png')
    
    # 將截圖轉換為Image開啟
    captcha_screenshot = Image.open('./captcha_screenshot.png')
    
    # 取得當前網頁大小
    window_size = driver.get_window_size()
    window_width = window_size['width']
    
    # 將截圖寬度除以網頁寬度，作為後續等比縮放裁切範圍的依據
    scaling_factor = captcha_screenshot.width / window_width
    
    print("裁切captcha螢幕截圖")
    # 透過locate_captcha定位截圖中captcha的位置並裁切存取
    captcha_image = captcha_screenshot.crop(locate_captcha(captcha_element, scroll_down, scaling_factor))
    captcha_image.save('./captcha_image.png')
    
    # 2captcha API KEY
    api_key = 'YOUR_API_KEY' 
    
    print("辨識captcha中")
    # 透過2captcha來自動辨識captcha
    solver = TwoCaptcha(api_key)
    
    # minLength = 5 設定文字最短長度
    result = solver.normal('./captcha_image.png', maxLength = 5)
    
    # 從result當中取得認證碼並轉成大寫
    captcha_code = result['code'].upper()
    
    # 檢查captcha_code格式，錯誤原因有可能來自api判斷錯誤以及截圖不完整，一旦發生就用未經裁切的原圖去偵測認證碼
    if len(captcha_code) < 5:
        print('Captcha格式錯誤')
        result = solver.normal('./captcha_screenshot.png', maxLength = 5)
        captcha_code = result['code'].upper()
        
    print(f"辨識完成，captcha答案是: {captcha_code}")    
    # 輸入: Captcha認證碼
    captcha_input = driver.find_element(By.XPATH, '//*[@id="captchaInput_captchaKey"]')
    captcha_input.send_keys(captcha_code)
    
    # 點擊: 搜尋
    search = driver.find_element(By.XPATH, '//*[@id="goSearch"]')
    search.click()
    return captcha_input, captcha_code

def get_district_data():
    # 設定driver路徑與相關網頁開啟參數
    # exec_path = './chromedriver.exe'    
    # service = Service(executable_path=exec_path)
    service = Service(ChromeDriverManager().install())
    options = webdriver.ChromeOptions()
    
    # 關閉以下Chrome內建優化以順利爬取
    # sandbox: 可防止惡意代碼對電腦造成傷害，但有時會影響selenium操作
    options.add_argument('--no-sandbox')
    # /dev/shm: 減少內存使用
    options.add_argument('--disable-dev-shm-usage')
    
    # 設定螢幕大小以防止在不同電腦運行時後續captcha截圖會跑掉
    options.add_argument('--window-size = 1200, 800')
    driver = webdriver.Chrome(service = service, options = options)
    
    # 開啟欲爬取的url
    url = 'https://www.ris.gov.tw/info-doorplate/app/doorplate/main?retrievalPath=%2Fdoorplate%2F'
    driver.get(url)
    
    # 設定智能等待時間
    driver.implicitly_wait(5)
    
    print("以編訂日期、編訂類別查詢")
    # 點擊: 以編訂日期、編訂類別查詢
    query_selection = driver.find_element(By.XPATH, '/html/body/div/div[4]/div/div/form/div[1]/fieldset/div/div[1]/button')
    query_selection.click()
    
    print("選取: 臺北市")
    # 點擊: 臺北市
    city_selection = driver.find_element(By.XPATH, '//*[@id="mapForm"]/div/fieldset/div[3]/map/area[1]')
    city_selection.click()
    
    print("選取: 中山區")
    # 找出下拉選單並選擇中山區
    district_dropdown = driver.find_element(By.XPATH, '//*[@id="areaCode"]')
    Select(district_dropdown).select_by_visible_text('中山區')
    
    print("選取: 起迄日期")
    # 點擊: 起始日期
    from_ = driver.find_element(By.XPATH, '//*[@id="sDate"]')
    from_.click()
    
    # 找出月份下拉選單後選擇一月
    from_month_dropdown = driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div/div/select[1]')
    Select(from_month_dropdown).select_by_visible_text('一月')
    
    # 找出年份下拉選單選擇民國111年
    from_year_dropdown = driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div/div/select[2]')
    Select(from_year_dropdown).select_by_visible_text('民國111年')
    
    # 點擊: 民國111年一月一日 
    from_date = driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/table/tbody/tr[1]/td[7]/a')
    from_date.click()
    
    # 點擊: 結束日期
    to_ = driver.find_element(By.XPATH, '//*[@id="eDate"]')
    to_.click()
    
    # 點擊: 今日
    # 透過Class name的特殊名稱設計定位今日
    to_today = driver.find_element(By.CLASS_NAME, 'ui-state-default.ui-state-highlight')
    to_today.click()
    
    # 透過while設立重跑機制
    captcha_not_pass = True
    rerun = 0
    while captcha_not_pass:
        if rerun < 3:
            rerun += 1
            captcha_input, captcha_code = captcha_solver(driver)
            # 透過find elements檢查是否出現認證碼錯誤彈窗
            block = driver.find_elements(By.XPATH, '/html/body/div[4]/div/div[3]/button[1]')
            if block:
                print('Captcha輸入錯誤，請重新讀取')
                block[0].click()
                captcha_input.clear()
                
            else:
                print('通過Captcha!')
                captcha_not_pass = False
        else:
            captcha_not_pass = False
        
    # 建立list來儲存所需欄位
    doorplate_ls = []
    date_ls = []
    category_ls = []
    pages = int(driver.find_element(By.XPATH, '//*[@id="sp_1_result-pager"]').text)
    print(f"總共有{pages}頁")
    
    for i in range(pages):
        # 解決換頁後元素偶爾讀取失效
        try:
            print(f"正在進行第{i+1}頁...")
            time.sleep(0.5)
            # 建立temp_ls
            temp_doorplate = [element.text for element in driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="門牌資料"]')]
            temp_date = [element.text for element in  driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="編釘日期 "]')]
            temp_category = [element.text for element in driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="編釘類別  "]')]
        except:
            print('element失效，重新擷取')
            time.sleep(1)
            temp_doorplate = [element.text for element in driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="門牌資料"]')]
            temp_date = [element.text for element in  driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="編釘日期 "]')]
            temp_category = [element.text for element in driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="編釘類別  "]')]
            
        doorplate_ls.extend(temp_doorplate)
        date_ls.extend(temp_date)
        category_ls.extend(temp_category)
        
        # 如果是最後一頁則不要點擊下一頁
        if i != (pages-1):
            next_page = driver.find_element(By.XPATH, '//*[@id="next_result-pager"]/span')
            next_page.click()
    print(f"爬取完畢，一共蒐集到{len(doorplate_ls)}筆結果")
    
    # 將前面蒐集好資料的list組成dataframe並輸出
    output = pd.DataFrame({'門牌資料': doorplate_ls, '編訂日期': date_ls, '編訂類別': category_ls})
    output.head()
    output.to_csv('test1.csv', encoding = 'utf-8-sig', index = False)    
    driver.quit()
    return print('輸出完畢')

if __name__ == "__main__":
    get_district_data()
    
# doorplate =  driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="門牌資料"]')
# date = driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="編釘日期 "]')
# category = driver.find_elements(By.CSS_SELECTOR,'[data-jqlabel="編釘類別  "]')