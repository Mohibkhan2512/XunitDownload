# import requests
# from bs4 import BeautifulSoup
# import pandas as pd
# import time

# # # URL of the web page you want to scrape
# # # url2 = 'https://axiom.qualcomm.com/#/planner/playlists/342922/46'

# url = 'https://qclogin.qualcomm.com/siteminderagent/nocert/1728277761/smgetcred.scc?TYPE=16777217&REALM=-SM-QUALCOMM%20Intranet%20--%20Corporate%20IT%20[10%3a39%3a21%3a491]&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-DSOzBUwcsyVOqg7U7P2MnKGqDP8zvogQXX1qiTTd6ZFjEjC1r94%2bxS8ql3FvmtYn&TARGET=-SM-HTTPS%3a%2f%2faxiom%2equalcomm%2ecom%2f#/planner/playlists/342922/46'

# # url = 'https://qclogin.qualcomm.com/siteminderagent/nocert/1728277761/smgetcred.scc?TYPE=16777217&REALM=-SM-QUALCOMM%20Intranet%20--%20Corporate%20IT%20[10%3a39%3a21%3a491]&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-DSOzBUwcsyVOqg7U7P2MnKGqDP8zvogQXX1qiTTd6ZFjEjC1r94%2bxS8ql3FvmtYn&TARGET=-SM-HTTPS%3a%2f%2f'


# # # Your credentials
# username = 'pmohibkh'
# password = 'Son1c@2512'

# # # Send a GET request to the web page with authentication
# response = requests.get(url, auth=(username, password), verify=False)

# time.sleep(10)

# # response = requests.get(url, verify=False)
# print(f'response => {response.text}')

# # print(f'response => {response}')


# # Parse the HTML content of the page
# soup = BeautifulSoup(response.content, 'html.parser')

# # Find the table in the HTML
# table = soup.find('table')

# print(f'table => {table}')

# # Extract table headers
# headers = [header.text for header in table.find_all('th')]

# # Extract table rows
# rows = []
# for row in table.find_all('tr'):
#     cells = row.find_all('td')
#     if len(cells) > 0:
#         rows.append([cell.text for cell in cells])

# # Create a DataFrame from the extracted data
# df = pd.DataFrame(rows, columns=headers)

# # Display the DataFrame
# print(df)


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Set up the WebDriver
# driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver = webdriver.Edge()

axiomUrl = 'https://axiom.qualcomm.com/#/planner/playlists/405303/0'

# axiomUrl = 'https://axiom.qualcomm.com/#/planner/playlists/417833'

axiomUrl = 'https://axiom.qualcomm.com/#/reports/job/27310091'

axiomUrl = 'https://axiom.qualcomm.com/#/reports/job/27218829'

axiomUrl = 'https://axiom.qualcomm.com/#/planner/playlists/304583/49'

axiomUrl = 'https://axiom.qualcomm.com/#/planner/playlists/299686/145'

# axiom%2equalcomm%2ecom%2f#/planner/playlists/405303/0'

axiomUrlIntr = axiomUrl.strip('https://').split('.')
axiomUrlIntr = ('%2e').join(axiomUrlIntr)
axiomUrl = axiomUrlIntr.replace('/#', '%2f#')

print(f'axiomUrl => {axiomUrl}')

url = 'https://qclogin.qualcomm.com/siteminderagent/nocert/1728277761/smgetcred.scc?TYPE=16777217&REALM=-SM-QUALCOMM%20Intranet%20--%20Corporate%20IT%20[10%3a39%3a21%3a491]&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-DSOzBUwcsyVOqg7U7P2MnKGqDP8zvogQXX1qiTTd6ZFjEjC1r94%2bxS8ql3FvmtYn&TARGET=-SM-HTTPS%3a%2f%2faxiom%2equalcomm%2ecom%2f#/planner/playlists/342922/46'

url = 'https://qclogin.qualcomm.com/siteminderagent/nocert/1728284238/smgetcred.scc?TYPE=16777217&REALM=-SM-QUALCOMM%20Intranet%20--%20Corporate%20IT%20[12%3a27%3a18%3a9668]&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-DSOzBUwcsyVOqg7U7P2MnKGqDP8zvogQXX1qiTTd6ZFjEjC1r94%2bxS8ql3FvmtYn&TARGET=-SM-HTTPS%3a%2f%2faxiom%2equalcomm%2ecom%2f#/planner/playlists/405303/0'

url = 'https://qclogin.qualcomm.com/siteminderagent/nocert/1728285147/smgetcred.scc?TYPE=16777217&REALM=-SM-QUALCOMM%20Intranet%20--%20Corporate%20IT%20[12%3a42%3a27%3a3774]&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-DSOzBUwcsyVOqg7U7P2MnKGqDP8zvogQXX1qiTTd6ZFjEjC1r94%2bxS8ql3FvmtYn&TARGET=-SM-HTTPS%3a%2f%2faxiom%2equalcomm%2ecom%2f#/reports/job/27214331'

url = 'https://qclogin.qualcomm.com/siteminderagent/nocert/1728285147/smgetcred.scc?TYPE=16777217&REALM=-SM-QUALCOMM%20Intranet%20--%20Corporate%20IT%20[12%3a42%3a27%3a3774]&SMAUTHREASON=0&METHOD=GET&SMAGENTNAME=-SM-DSOzBUwcsyVOqg7U7P2MnKGqDP8zvogQXX1qiTTd6ZFjEjC1r94%2bxS8ql3FvmtYn&TARGET=-SM-HTTPS%3a%2f%2f' + axiomUrl



# Your credentials
username = 'pmohibkh'
password = 'Son1c@2512'


begin = time.time() 

# Open the web page
driver.get(url)

# Increase the wait time
wait = WebDriverWait(driver, 120)
import pyautogui
import time

time.sleep(1)

pyautogui.write(username)
pyautogui.press('tab')
pyautogui.write(password)
pyautogui.press('tab')

pyautogui.press('enter')

# Wait for the page to load completely
wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

# Execute JavaScript to ensure all tables are loaded
driver.execute_script("return document.readyState == 'complete'")


if 'job' in url:
    extractFromJob = True
else:
    extractFromJob = False

# time.sleep(5)

# tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))

# tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))

if extractFromJob == True:
    # tablesBtn = driver.find_element(By.CLASS_NAME, 'p-ripple p-element p-panel-header-icon p-panel-toggler p-link ng-tns-c2655002852-35 ng-star-inserted')
    tablesBtn = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'panelWrap.collapsed.ng-star-inserted.collapsed')))
    # 'panelWrap ng-star-inserted selected collapsed'pmohibkh   Son1c@2512  
    tablesBtn = tablesBtn[0]
    print(f'tablesBtn -> {tablesBtn} ')
    tablesBtn.click()

    tables = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'p-datatable-table.p-datatable-scrollable-table.p-datatable-resizable-table.ng-star-inserted')))

    # Select the first table in case of job link
    table = tables[0]
    extractFromJob = True

    # Extract table headers from thead
    headers = []
    thead = table.find_element(By.TAG_NAME, 'thead')
    for row in thead.find_elements(By.TAG_NAME, 'tr'):
        # print(f'tr => {row.text}')
        for header in row.find_elements(By.TAG_NAME, 'th'):
            headers.append(header.text.strip())

    print(f'headers => {headers}')

    # Extract table rows from tbody
    rows = []
    tbody = table.find_element(By.TAG_NAME, 'tbody')
    for row in tbody.find_elements(By.TAG_NAME, 'tr'):
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) > 0:
            rows.append([cell.text.strip() for cell in cells])

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(rows, columns=headers)

    # Specify the headers you want to extract
    specific_headers = ['unfold_more\nTest Case', 'unfold_more\nResults']  # Replace with the actual headers you want

    # Filter the DataFrame to include only these headers
    df_filtered = df[specific_headers]

    print(df_filtered)

    # Close the WebDriver
    driver.quit()

else:
    tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))
    table = tables[1]

    # Extract table headers from thead
    headers = []
    thead = table.find_element(By.TAG_NAME, 'thead')
    for row in thead.find_elements(By.TAG_NAME, 'tr'):
        # print(f'tr => {row.text}')
        for header in row.find_elements(By.TAG_NAME, 'th'):
            headers.append(header.text.strip())

    print(f'headers => {headers}')

    # Extract table rows from tbody
    rows = []
    tbody = table.find_element(By.TAG_NAME, 'tbody')
    for row in tbody.find_elements(By.TAG_NAME, 'tr'):
        cells = row.find_elements(By.TAG_NAME, 'td')
        if len(cells) > 0:
            rows.append([cell.text.strip() for cell in cells])

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(rows, columns=headers)

    # Specify the headers you want to extract
    specific_headers = ['unfold_more\nName', 'Parameters']  # Replace with the actual headers you want

    # Filter the DataFrame to include only these headers
    df_filtered = df[specific_headers]

    # print(df_filtered['Parameters'])

    # parameters = df_filtered['Parameters'].to_string()
    # parameters = df_filtered['Parameters'].astype(str).tolist()
    # parameters = parameters.split(',')

    parametersList = []

    for value in df_filtered['Parameters']:
    # for index, row in df_filtered['Parameters'].iterrows():
        parametersList.append([value])

    # print(f'parameters => {parametersList}')

    # stringData = parameters[0]
    # stringData = ''

    # for data in parameters:
    #     stringData += data

    # stringData = stringData.split(', ')

    # print(f'stringData => {stringData}')

    testParams = []

    for row in parametersList:
        row = row[0].split(', ')
        paramData = ''
        for data in row:
            if 'noDeviceReset=' in data:
                data = data.split('=')
                # print(f'data => {data}')
                data = data[1]
                paramData += f'--noDeviceReset {data} '
            elif 'paramEntry=' in data:
                data = data.split('=')
                # print(f'data => {data}')
                data = data[1]
                paramData += f'--paramEntry {data} '
            elif 'target=' in data:
                data = data.split('=')
                # print(f'data => {data}')
                data = data[1]
                paramData += f'--target {data} '
            elif 'testName=' in data:
                data = data.split('=')
                # print(f'data => {data}')
                data = data[1]
                paramData += f'--testName {data} '
            elif 'SPName=' in data:
                data = data.split('=')
                # print(f'data => {data}')
                data = data[1]
                paramData += f'--SPName {data} '

        testParams.append([paramData])
    
    # print(f'testParams => {testParams}')
    
    print(f'len( testParams => {len(testParams)}')

    print(len(df_filtered['unfold_more\nName']))

    testNameWithParams = []
    testNameList = df_filtered['unfold_more\nName'].astype(str).tolist()

    for index, testName in enumerate(testNameList):
        testNameWithParams.append([testName, testParams[index]])
    
    
    print(f'testNameWithParams => {testNameWithParams}')

    # Close the WebDriver
    driver.quit()

end = time.time()

print(f"Total runtime of the program is {end - begin}") 
