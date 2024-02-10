import openpyxl
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl.styles import Font

driver = webdriver.Chrome(ChromeDriverManager().install())
driver.get('https://iotvega.com/product')

scroll_content = driver.find_element(By.CLASS_NAME, 'product-name')
scroll_content.click()

title = driver.find_element(By.XPATH, "/html/body/section[1]/div/div/div/div/div/h1").text # title
elements = driver.find_element(By.XPATH, "/html/body/section[3]/div/div/div[1]/table/tbody").text # all elements
some_el = driver.find_element(By.XPATH, "/html/body/section[3]/div/div/div[1]/table/tbody/tr[1]/td[1]").text # name of characteristic
value = driver.find_element(By.XPATH, "/html/body/section[3]/div/div/div[1]/table/tbody/tr[1]/td[2]").text # value

len_list_of_elements = len(elements.split('\n')) * 2

workbook = openpyxl.Workbook()
sheet = workbook.active

# Ширина столбцов
sheet.column_dimensions['A'].width = 52
sheet.column_dimensions['B'].width = 35

# Определил строки A1 и A2 под заголовки
sheet["A1"] = 'Характеристика'
sheet["A1"].font = Font(bold=True)
sheet["A2"] = title
sheet["A2"].font = Font(bold=True)

# Цикл для записи значений в ячейки
j = 3
while j < len(elements.split('\n')) + 2:
    # Цикл для получения значений из XPATH
    for i in range(1, len_list_of_elements, 2):

        some_el = driver.find_element(By.XPATH, f'/html/body/section[3]/div/div/div[1]/table/tbody/tr[{i}]/td[1]').text
        value = driver.find_element(By.XPATH, f'/html/body/section[3]/div/div/div[1]/table/tbody/tr[{i}]/td[2]').text

        sheet['A' + str(j)].value = some_el
        sheet['B' + str(j)].value = value

        j += 1

workbook.save("results.xlsx")
