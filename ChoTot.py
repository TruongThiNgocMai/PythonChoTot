from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import os, inspect
import pandas as pd
from pandas import ExcelWriter
from time import sleep
import re
import unidecode

CurDir = os.path.dirname(os.path.realpath(__file__))
chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 2}
chrome_options.add_experimental_option("prefs",prefs)

#Xóa những task cũ
def Clear():
   os.system("taskkill /f /im chromedriver.exe")
   os.system("taskkill /f /im chrome.exe")

#Mở excel sử dụng thư viện pandas
def Read_Excel(path,sheet):
   df = pd.read_excel(open(path,'rb'), sheet_name = sheet, dtype = object)
   return df

#Mở trình duyệt chromedriver sử dụng selenium
def Open_Browser():
   driver = webdriver.Chrome(CurDir+"\\chromedriver.exe")
   return driver

# click vào xpath nếu tồn tại
def check_exists_by_xpath(xpath,driver):
   while True:
      try:
         driver.find_element_by_xpath(xpath).click()
         break
      except:
         sleep(1)
         pass

#hàm nhập
def nhap():
   intCheck = 0
   dfTP = pd.read_excel(CurDir+"\\Input\\TinhThanh.xlsx")
   df = pd.DataFrame(dfTP) 
   while True:
      if intCheck  == 1:
         break
      strTenTinh = input("Nhập tỉnh thành bạn muốn tìm kiếm (viết hoa có dấu ):   ")
      for i in df.index:  
         if (df["Tỉnh Thành"][i]) != str(strTenTinh):
               # print("Gõ đúng cái coi! Bực cả mình")
               pass
         elif (df["Tỉnh Thành"][i]) == str(strTenTinh):
               strTenTinh = df["Tỉnh Thành"][i]
               intCheck = 1
   return strTenTinh

def to_slug(strName):
   strName = unidecode.unidecode(strName).lower()
   return re.sub(r'[\W_]+', '-', strName)

def get_Value(xpath, driver):
   try:
      value = driver.find_element_by_xpath(xpath)
      value = value.text
      return value
   except:
      pass
   

#Hàm xử lý chính
def Process_Main():
   writer = pd.ExcelWriter(CurDir+'\\Output\\Output.xlsx')
   dfTTDN = pd.DataFrame(columns=['STT','Tin tức','Diện tích','Giá phòng (tháng)','SDT liên hệ','Địa chỉ BĐS','Thông tin thêm','Đường dẫn'])
   strTenTinh = nhap()

   # mở tới các trang
   driver = webdriver.Chrome(CurDir+"\\chromedriver.exe",chrome_options=chrome_options)
   driver.maximize_window() #max win 
   driver.get("https://www.chotot.com/")
   sleep(2) 
   driver.find_element_by_xpath('//*[@id="boxListCate"]/ul/li[1]/a/div').click() #vào bất động sản
   sleep(2) 
   check_exists_by_xpath('//*[@id="tooltip_btn_save_search"]/div[4]',driver) # tắt popup nếu có
   sleep(2)
   driver.find_element_by_xpath('//*[@id="app"]/div[2]/main/div/div[1]/div[2]/div[3]/div/div[1]/div/a[2]').click() #vào cho thuê và toàn quốc
   driver.find_element_by_xpath('//*[@id="regionRef"]/div').click() # chọn vùng miền 
   sleep(2) 
   intCheck = 0 
   
   listTenTinh = driver.find_elements_by_xpath("//*[@id='regionRef']/div[2]/div/ul/li[*]/div/a")
   for TenTinh in listTenTinh:
      if intCheck == 1:
         break 
      else:
         if str(TenTinh.text) == strTenTinh:
            driver.get(TenTinh.get_attribute("href"))
            sleep(2)
            driver.find_element_by_xpath('//*[@id="categoryRef"]/div/div/span').click() #tất cả
            intCheck = 1
            element = driver.find_elements_by_xpath('//*[@id="categoryRef"]/div[2]/div/ul/li[*]/div/a') #lấy phòng trọ
            count = 0
   for option in element:
      if str(option.text) == "Phòng trọ":
         driver.get(option.get_attribute("href"))
         driver.get('https://nha.chotot.com/'+to_slug(strTenTinh)+'/thue-phong-tro?price=0-3000000') 
         sleep(2)
         intI = 1
         count = 0
         while intI <= 1:
            driver.get('https://nha.chotot.com/'+to_slug(strTenTinh)+'/thue-phong-tro?page='+str(intI)+'&price=0-3000000')
            intI += 1
            sleep(2)
            list_URLS = driver.find_elements_by_xpath("//*[@id='app']/div[2]/main/div/div[1]/div[2]/main/div/div[1]/div[6]/div/div[2]/ul/div[*]/li/a")
            for idx_url, i_url in enumerate(list_URLS):
               driver1 = Open_Browser()
               driver1.maximize_window()
               driver1.get(i_url.get_attribute("href"))
               try:
                  check_exists_by_xpath('//*[@id="tooltip_btn_save_ad"]/div[3]',driver1)
               except:
                  pass
               check_exists_by_xpath('//*[@id="app"]/div[2]/main/article/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[3]',driver1)
               check_exists_by_xpath('//*[@id="app"]/div[2]/main/article/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div[1]/div/div/div/img',driver1)
               listSDT = get_Value('//*[@id="app"]/div[2]/main/article/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div[1]/div/div/div/span', driver1)
               # listDienTich = get_Value("//*[@id='app']/div[2]/main/article/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/span[2]",driver1)
               listTinTuc = get_Value("//*[@id='app']/div[2]/main/article/div[1]/div[2]/div[1]/div[2]/h1",driver1)
               listGia = get_Value("//*[@id='app']/div[2]/main/article/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/div[1]/span[1]/span[1]",driver1)
               listDiaChi = get_Value("//*[@id='app']/div[2]/main/article/div[1]/div[2]/div[1]/div[2]/div[3]/div/div[2]/div",driver1)
               listThongTinThem = get_Value("//*[@id='app']/div[2]/main/article/div[1]/div[2]/div[1]/div[2]/p",driver1)
               count+=1
               dfTTDN.loc[count,'STT']=str(count)
               dfTTDN.loc[count,'Tin tức'] = str(listTinTuc)
               # dfTTDN.loc[count,'Diện tích'] = str(listDienTich.replace('- ',''))
               dfTTDN.loc[count,'Giá phòng (tháng)'] = str(listGia)
               dfTTDN.loc[count,'SDT liên hệ'] = str(listSDT)
               dfTTDN.loc[count,'Địa chỉ BĐS'] = str(listDiaChi)
               dfTTDN.loc[count,'Thông tin thêm'] = str(listThongTinThem)
               
               dfTTDN.loc[count,'Đường dẫn'] = str(i_url.get_attribute("href"))
               driver1.quit()
      dfTTDN.to_excel(writer, sheet_name='Sheet1', index=False)
   writer.save()
   driver.quit()
#Phần thực thi
if __name__ == "__main__":  # bắt đâu chạy từ đây
   print ("Bắt đầu quy trình")
   Clear()
   Process_Main()
   print ("Kết thúc quy trình")