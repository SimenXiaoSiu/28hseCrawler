import time
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Tools.scripts.fixdiv import report

def get_data():
    result_list = []
    temp_list = []
    url = 'https://www.28hse.com/buy'
    #url = 'https://data.28hse.com/kl'
    browser = webdriver.Chrome("G:/WebDriver/chromedriver.exe")
    browser.get(url)
    try:
        temp_list = ['District','Cat','Reported Size','Actual Size','$','LandLord']
        result_list.append(temp_list.copy())
        temp_list.clear()

        WebDriverWait(browser, 3).until(
            EC.presence_of_element_located((By.CLASS_NAME,"searching_box2_inner"))
            )
        browser.find_element_by_link_text("新界").click()
        time.sleep(1)
        browser.find_element_by_link_text("住宅").click()
        time.sleep(1)
        #browser.find_element_by_link_text("200-400萬").click()
        #browser.find_element_by_link_text("400-800萬").click()
        #time.sleep(1)
        browser.find_element_by_link_text("最新").click()
        time.sleep(1)
        for x in range(20):
            WebDriverWait(browser, 10).until(
                EC.presence_of_element_located((By.ID,"search_main_div"))
                )
            time.sleep(2)
            for item in browser.find_elements_by_class_name('agentad_ul'):
                #district = item.find_element_by_xpath('.//a[@data-district-id]').text
                #cat = item.find_element_by_xpath('.//a[@data-cat-id]').text
                #report_size = item.find_element_by_xpath(".//*[contains(text(), '建築面積')]").text
                #actual_size = item.find_element_by_xpath(".//*[contains(text(), '實用面積')]").text
                #sold = item.find_element_by_xpath('.//div[@class="price_class"]').text
                #land_lord = item.find_element_by_xpath('.//div[@class="landlord_2"]').text
                #temp_list.insert(len(temp_list), district)
                #temp_list.insert(len(temp_list), cat)
                #temp_list.insert(len(temp_list), report_size)
                #temp_list.insert(len(temp_list), actual_size)
                #temp_list.insert(len(temp_list), sold)
                #temp_list.insert(len(temp_list), land_lord)
                #result_list.append(temp_list.copy())
                #temp_list.clear()
                if len(item.find_elements_by_xpath('.//a[@data-district-id]')) > 0:
                    district = item.find_element_by_xpath('.//a[@data-district-id]').text
                    temp_list.insert(len(temp_list), district)
                else:
                    temp_list.insert(len(temp_list), "NA")
                    
                if len(item.find_elements_by_xpath('.//a[@data-cat-id]')) > 0:
                    cat = item.find_element_by_xpath('.//a[@data-cat-id]').text
                    temp_list.insert(len(temp_list), cat)
                else:
                    temp_list.insert(len(temp_list), "NA")
                    
                if len(item.find_elements_by_xpath(".//*[contains(text(), '建築面積')]")) > 0:
                    report_size = item.find_element_by_xpath(".//*[contains(text(), '建築面積')]").text
                    temp_list.insert(len(temp_list), report_size)
                else:
                    temp_list.insert(len(temp_list), "NA")
                    
                if len(item.find_elements_by_xpath(".//*[contains(text(), '實用面積')]")) > 0:
                    actual_size = item.find_element_by_xpath(".//*[contains(text(), '實用面積')]").text
                    temp_list.insert(len(temp_list), actual_size)
                else:
                    temp_list.insert(len(temp_list), "NA")
                    
                if len(item.find_elements_by_xpath('.//div[@class="price_class"]')) > 0:
                    sold = item.find_element_by_xpath('.//div[@class="price_class"]').text
                    temp_list.insert(len(temp_list), sold)
                else:
                    temp_list.insert(len(temp_list), "NA")
                    
                if len(item.find_elements_by_xpath('.//div[@class="landlord_2"]')) > 0:
                    land_lord = item.find_element_by_xpath('.//div[@class="landlord_2"]').text
                    temp_list.insert(len(temp_list), land_lord)
                else:
                    temp_list.insert(len(temp_list), "NA")
                    
                result_list.append(temp_list.copy())
                temp_list.clear()            
            browser.find_element_by_link_text("次頁").click()
            time.sleep(1)

    finally:
        browser.quit()
        
    return result_list

def write_to_excel(list):
    wb = Workbook()

    # grab the active worksheet
    ws = wb.active
    
    for i in list:
        ws.append(i)    
    
    # Save the file


if __name__ == '__main__':
    result_list = get_data()
    write_to_excel(result_list)
