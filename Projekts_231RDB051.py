from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import Workbook, load_workbook 
from selenium.webdriver.remote.webelement import WebElement

# Lietotājs ievada lietotājvārdu un paroli
login = input("Lūdzu, ievadiet Lietotājvārds: ")
password = input("Lūdzu, ievadiet Parole: ")

# Iestatījumi Selenium draivera izveidei
service = Service()
option = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=option)

# Ielādē mājaslapu, kur notiks ienākšana
url = "https://id2.rtu.lv/openam/UI/Login?module=LDAP&locale=lv"
driver.get(url)

# Funkcija elementa noklikšķināšanai
def elemnt_click(tag_name, attribute, attribute_value):
    link = driver.find_element(By.XPATH, f'//{tag_name}[@{attribute}="{attribute_value}"]')
    link.click()

# Atrast ievades laukus un ievadīt lietotājvārdu un paroli
find = driver.find_element(By.ID, "IDToken1")
find.send_keys(login)
find = driver.find_element(By.ID, "IDToken2")
find.send_keys(password)

# Iesniegt formu
find.submit()

# Noklikšķināt uz "Studentiem" saites
elemnt_click("a", "title", "Studentiem")

# Definēt funkciju priekšmeta datu izgūšanai
def pars_prieksmets(id):

    # Izveidot Excel darba grāmatu un saiti uz aktīvo lapu
    myworkbook=Workbook()
    worksheet= myworkbook.active

    # Iegūt priekšmeta nosaukumu un izgriezt nepieciešamo informāciju
    vards = driver.find_element(By.XPATH, f'//a[@href="https://estudijas.rtu.lv/course/view.php?id={id}"]').text
    vards = vards[:vards.index("(")].strip()

    # Noklikšķināt uz priekšmeta saites un atvērt vērtējumu lapu
    elemnt_click("a", "href", f'https://estudijas.rtu.lv/course/view.php?id={id}')
    elemnt_click("a", "id", "action-menu-toggle-2")
    elemnt_click("a", "href", f"https://estudijas.rtu.lv/grade/report/user/index.php?id={id}")

    # Atrodoties uz vērtējumu lapas, iegūt tabulu un visas rindiņas
    table = driver.find_element(By.TAG_NAME, "table")
    allRows = table.find_elements(By.TAG_NAME, "tr")

    index = 1
    for  row in allRows :
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) >= 4:
            
            # Izprintēt studenta vārdu un vērtējumu
            if len(cells[0].find_elements(By.TAG_NAME, "a")) > 0:
                print(f'{cells[0].find_element(By.TAG_NAME, "a").text} - {cells[2].text}')
                worksheet[f'A{index}'] = cells[0].find_element(By.TAG_NAME, "a").text
                worksheet[f'B{index}'] = cells[2].text
                index += 1

    # Saglabāt rezultātus Excel failā
    myworkbook.save(f'ocenki_{vards}.xlsx')
    myworkbook.close()

    # Atgriezties atpakaļ uz iepriekšējo lapu
    driver.back()
    driver.back()

# Atrast gada elementu un visus priekšmetu saites
gads = driver.find_element(By.ID, "Pluto_48_u108l1n154138_127827_group2")
prieksmeti = gads.find_elements(By.TAG_NAME, "table")
for prieksmets in prieksmeti:
    link = prieksmets.find_element(By.TAG_NAME, "a")
    code = link.get_dom_attribute("href")[-6:]
    pars_prieksmets(code)

# Pauze, lai pietiktu laika redzēt rezultātus pirms lapas aizvēršanas
time.sleep(10)

# Aizvērt pārlūku
driver.quit()
