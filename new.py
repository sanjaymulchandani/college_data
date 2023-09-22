from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl


file_path = 'data_extract.xlsx'
workbook = openpyxl.load_workbook(file_path)
active_sheet= workbook.active
print(active_sheet)

# Xpath

google_search = 'list of universities in india wikipedia'
google_search_xpath = '//*[@id="APjFqb"]'
wiki_link_xpath = "//h3[normalize-space()='List of universities in India']"
main_page_title = 'List of universities in India - Wikipedia'
table_xpath = "//table[@class='wikitable sortable jquery-tablesorter']"
table_header_xpath = '//*[@id="mw-content-text"]/div[1]/table/thead/tr/th[1]'
list_1 = '//*[@id="mw-content-text"]/div[1]/table/tbody/tr[1]/td[1]/a[2]'


driver = webdriver.Chrome()
actions = ActionChains(driver)
url = 'https://www.google.com'
driver.get(url)
driver.maximize_window()

def find_college_list():
    # Wait until google search button is loaded
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, google_search_xpath)))

    # locate search input box and perform search
    driver.find_element(By.XPATH, google_search_xpath).send_keys(google_search)
    driver.find_element(By.XPATH, google_search_xpath).send_keys(Keys.ENTER)
    
    # Wait until wikipedia link is loaded
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, wiki_link_xpath)))

    # click on wiki link of list of universities in India
    driver.find_element(By.XPATH, wiki_link_xpath).click()
    
    # Wait until list of colleges table is loaded
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, table_xpath)))

    # once table is loaded scroll to table area
    table_obj = driver.find_element(By.XPATH, table_xpath)
    driver.execute_script("arguments[0].scrollIntoView();", table_obj)

find_college_list()
time.sleep(5)


"""# Extract table data
def Andhra_Pradesh():
    link_Andhra_Pradesh = "//a[@title='List of educational institutions in Andhra Pradesh']"
    driver.find_element(By.XPATH, link_Andhra_Pradesh).click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="footer-places-about"]/a')))

Andhra_Pradesh()

# Central instutions
row_xpath = '//*[@id="mw-content-text"]/div[1]/table[2]/thead/tr/th'
row_ele = driver.find_elements(By.XPATH, row_xpath)
col_xpath = '//*[@id="mw-content-text"]/div[1]/table[2]/tbody/tr/td'
col_ele = driver.find_elements(By.XPATH, col_xpath)

print("ROWS: ", len(row_ele))
print("COLS: ",len(col_ele))
"""

def central_uni():
    central_uni_link_xpath = "//a[normalize-space()='Centraluniversities']"
    table_ele_xpath = "//th[@title='Sort ascending'][normalize-space()='University']"

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, central_uni_link_xpath)))

    driver.find_element(By.XPATH, central_uni_link_xpath).click()

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, table_ele_xpath)))

    col_xpath = '//body[1]/div[2]/div[1]/div[3]/main[1]/div[3]/div[3]/div[1]/table[2]/tbody[1]/tr[1]/td'
    col_ele = driver.find_elements(By.XPATH, col_xpath)

    row_xpath = '//body[1]/div[2]/div[1]/div[3]/main[1]/div[3]/div[3]/div[1]/table[2]/tbody[1]/tr/td[1]'
    row_ele = driver.find_elements(By.XPATH, row_xpath)

    print("ROW: ",len(row_ele))
    print("COL: ",len(col_ele))

    for row in range(1, len(row_ele)+1):
        for col in range(1, len(col_ele)):
            data_xpath = '//body[1]/div[2]/div[1]/div[3]/main[1]/div[3]/div[3]/div[1]/table[2]/tbody[1]/tr[{}]/td[{}]'.format(row, col)
            data = driver.find_element(By.XPATH, data_xpath)
            print(data.text)
            active_sheet.cell(row=row, column=col).value = data.text
        print('\n')
    workbook.save(file_path)
    workbook.close()








# Xpath_state_uni
state_uni_link_xpath = "//a[normalize-space()='Stateuniversities']"
andhra_pradesh_xpath = "//span[@id='Andhra_Pradesh']"
assam_xpath = "//span[@id='Assam']"
bihar_xpath = "//span[@id='Bihar']"
chandigrah_xpath = "//span[@id='Chandigarh']"
chhattisgarh_xpath = "//span[@id='Chhattisgarh']"
delhi_xpath = "//span[@id='Delhi']"
goa_xpath = "//span[@id='Goa']"
gujarat_xpath = "//span[@id='Gujarat']"
haryana_xpath = "//span[@id='Haryana']"
himachal_xpath = "//span[@id='Himachal_Pradesh']"
jammu_kashmir_xpath = "//span[@id='Jammu_and_Kashmir']"
jharkand_xpath = "//span[@id='Jharkhand']"
karnataka_xpath = "//span[@id='Karnataka']"
kerala_xpath = "//span[@id='Kerala']"
madhya_pradesh_xpath = "//span[@id='Madhya_Pradesh']"
maharashtra_xpath = "//span[@id='Maharashtra']"
manipur_xpath = "//span[@id='Manipur']"
odisha_xpath = "//span[@id='Odisha']"
punjab_xpath = "//span[@id='Punjab']"
rajasthan_xpath = "//span[@id='Rajasthan']"
tamil_nadu_xpath = "//span[@id='Tamil_Nadu']"
telangana_xpath = "//span[@id='Telangana']"
tripura_xpath = "//span[@id='Tripura']"
uttar_pradesh_xpath = "//span[@id='Uttar_Pradesh']"
uttrakhand_xpath = "//span[@id='Uttar_Pradesh']"
west_bengal_xpath = "//span[@id='West_Bengal']"



WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, state_uni_link_xpath)))

driver.find_element(By.XPATH, state_uni_link_xpath).click()

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//a[normalize-space()='Education in India']")))

table_obj_1 = driver.find_element(By.XPATH, andhra_pradesh_xpath)
driver.execute_script("arguments[0].scrollIntoView();", table_obj_1)

col_xpath = '//body[1]/div[2]/div[1]/div[3]/main[1]/div[3]/div[3]/div[1]/table[4]/tbody[1]/tr[1]/td'
col_ele = driver.find_elements(By.XPATH, col_xpath)

row_xpath = '//body[1]/div[2]/div[1]/div[3]/main[1]/div[3]/div[3]/div[1]/table[4]/tbody[1]/tr/th[1]'
row_ele = driver.find_elements(By.XPATH, row_xpath)

print("ROW: ",len(row_ele))
print("COL: ",len(col_ele))


for row in range(1, len(row_ele)+1):
    for col in range(1, len(col_ele)):
        data_xpath = '//body[1]/div[2]/div[1]/div[3]/main[1]/div[3]/div[3]/div[1]/table[4]/tbody[1]/tr[{}]/td[{}]'.format(row, col)
        data = driver.find_element(By.XPATH, data_xpath)
        print(data.text)
    print('\n')

table_obj_2 = driver.find_element(By.XPATH, assam_xpath)
driver.execute_script("arguments[0].scrollIntoView();", table_obj_2)




