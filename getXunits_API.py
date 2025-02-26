from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
# from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
from pathlib import Path
import shutil

# 27320486
# 27310948
# 27301133
# 27290882
# 27289593
# 27289418
# 27289258
# 27287720
# 27287466
# 27240264
# 27238565
# 27221690
# 27219852
# 27212903
# 27189124
# 27185175

joblinks = '''
28008936
27993166
27988588
27984760
27965605
27945505
27941174
27940578
27884030
27883369
27881157
27877412
'''
storeXunitsAt = 'C:\Dropbox\Mohib\CBXunits\ES15'
begin = time.time()

for job in joblinks.split('\n'):
    if job != '':
        axiomUrl = 'https://axiom.qualcomm.com/#/reports/jobs/' + job
        print(f'axiomUrl => ', axiomUrl)
    else:
        continue
   
   # Set up the WebDriver
    driver = webdriver.Edge()

    url = axiomUrl.replace('reports', 'logs')

    print(f'url => {url}')


    def your_credentials():
        # Your credentials
        username = 'pmohibkh'
        password = 'Son1c@2512'

        return username, password

    username, password = your_credentials()

    try:
        # Open the web page
        driver.get(url)
    except:
        continue

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

        
    tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))
    # print(f'tables -> {tables}')
    table = tables[0]

    # Extract table headers from thead
    headers = []
    thead = table.find_element(By.TAG_NAME, 'thead')
    for row in thead.find_elements(By.TAG_NAME, 'tr'):
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

    # print('df => ', df)

    atagsList = df['unfold_more\nName'].astype(str).tolist()
    atagsListFinal = []

    print(f'Log folders for the job {job} => {atagsList}')

    saveDone = 1
    for index, tag in enumerate(atagsList):
        tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))
        print(f'tables -> {tables}')
        table = tables[0]

        tbody = table.find_element(By.TAG_NAME, 'tbody')
        for row in tbody.find_elements(By.TAG_NAME, 'tr'):
            atags = row.find_elements(By.TAG_NAME, 'a')
            # print(f'atags => {atags}')

            # atagsListFinal.append(atag)
            for atag in atags:
                if atag.text == tag:
                    curAtag = atag
                    break

        # print(f'index => {index}')
        # print(f'tag => {curAtag.text}')
        
        curAtag.click()
        PostprocessingFound = False
        # time.sleep(2)
        # # break
        atagsListForPostProssessing = []
        tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))
        # print(f'tables two -> {tables}')

        # print(f'headers two => {headers}')

    
        table = tables[0]

        tbody = table.find_element(By.TAG_NAME, 'tbody')
        for row in tbody.find_elements(By.TAG_NAME, 'tr'):
            atag = row.find_elements(By.TAG_NAME, 'a')
            atagsListForPostProssessing.append(atag)

        # print(f'atagsListForPostProssessing -> {atagsListForPostProssessing}')


        for i, atag in enumerate(atagsListForPostProssessing):
            if atag[0].text == 'Postprocessing':
                PostprocessingFound = True
                print(f'Postprocessing found for {tag}...')

                atag[0].click()

                # click on xunit folder

                atagsListForXunit = []
                tables = wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, 'table')))

                table = tables[0]

                tbody = table.find_element(By.TAG_NAME, 'tbody')
                for row in tbody.find_elements(By.TAG_NAME, 'tr'):
                    atagXunit = row.find_elements(By.TAG_NAME, 'a')
                    atagsListForXunit.append(atagXunit)

                print(f'atagsListForXunit => {atagsListForXunit}')

                for atagXunit in atagsListForXunit:
                    if atagXunit[0].text == 'XUnit_Summary-Genarate':
                        print(f'Postprocessing found in Postprocessing for {tag}...')
                        atagXunit[0].click()
                        break

                checkBoxDiv = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'p-checkbox.p-component')))
                checkBoxDiv[0].click()

                time.sleep(0.5)

                downloadBtn = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'p-element.standard.p-button.p-component.ng-star-inserted')))

                downloadBtn[1].click()
                print('downloadBtn clicked...')

                time.sleep(1)

                # Open the downloads page in Edge
                driver.get('edge://downloads/')

                # Wait for the downloads page to load
                time.sleep(2)

                # Wait for the "Keep" button to be clickable
                # save = 'save'+str(index+1)
                save = 'save'+str(saveDone)
                print(f'save state => {save}')
                keep_button = wait.until(EC.presence_of_all_elements_located((By.ID, save)))
                saveDone += 1

                print(f'keep_button {keep_button}')
                keep_button[0].click()
                # keep_button.click()
                print('keep_button clicked...')
                # except:
                #     continue

                bg = time.time()
                print(f'sleep begin {bg}')
                print('sleep for 10 sec')
                time.sleep(5)
                eg = time.time()
                print(f'sleep end {eg-bg}')
                # Define the download directory and the target directory

                # Get the user's home directory
                home_dir = str(Path.home())

                # Construct the path to the download directory
                if os.name == 'nt':  # For Windows
                    download_dir = os.path.join(home_dir, 'Downloads')

                print(f"Download directory: {download_dir}")

                # download_dir = '/path/to/download/directory'
                target_dir = storeXunitsAt

                # Get the list of files in the download directory
                files = os.listdir(download_dir)

                # Sort files by modification time in descending order
                files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)

                # Get the most recently downloaded file
                recent_file = files[0]

                print(f'recent_file => {recent_file}')

                if 'xunit_report' in recent_file:
                    # Construct full file paths
                    src_path = os.path.join(download_dir, recent_file)
                    dst_path = os.path.join(target_dir, recent_file)
                    # Move the file
                    shutil.move(src_path, dst_path)

                    print(f"Moved {recent_file} to {target_dir}")
                break
            
            elif atag[0].text == 'XUnit_Summary-Genarate':
                PostprocessingFound = True
                print(f'XUnit_Summary found for {tag}...')
                
                atag[0].click()

                checkBoxDiv = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'p-checkbox.p-component')))
                checkBoxDiv[0].click()

                time.sleep(0.5)

                downloadBtn = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'p-element.standard.p-button.p-component.ng-star-inserted')))

                downloadBtn[1].click()
                print('downloadBtn clicked...')

                # Open the downloads page in Edge
                driver.get('edge://downloads/')

                # Wait for the downloads page to load
                time.sleep(2)

                # Wait for the "Keep" button to be clickable
                save = 'save'+str(saveDone)
                print(f'save state => {save}')
                keep_button = wait.until(EC.presence_of_all_elements_located((By.ID, save)))
                saveDone += 1

                print(f'keep_button {keep_button}')
                keep_button[0].click()
                # keep_button.click()
                print('keep_button clicked...')

                bg = time.time()
                print(f'sleep begin {bg}')
                print('sleep for 10 sec')
                time.sleep(5)
                eg = time.time()
                print(f'sleep end {eg-bg}')
                # Define the download directory and the target directory

                # Get the user's home directory
                home_dir = str(Path.home())

                # Construct the path to the download directory
                if os.name == 'nt':  # For Windows
                    download_dir = os.path.join(home_dir, 'Downloads')

                print(f"Download directory: {download_dir}")

                # download_dir = '/path/to/download/directory'
                target_dir = storeXunitsAt

                # Get the list of files in the download directory
                files = os.listdir(download_dir)

                # Sort files by modification time in descending order
                files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)), reverse=True)

                # Get the most recently downloaded file
                recent_file = files[0]

                print(f'recent_file => {recent_file}')

                if 'xunit_report' in recent_file:
                    # Construct full file paths
                    src_path = os.path.join(download_dir, recent_file)
                    dst_path = os.path.join(target_dir, recent_file)
                    # Move the file
                    shutil.move(src_path, dst_path)

                    print(f"Moved {recent_file} to {target_dir}")
                break
        
        if PostprocessingFound == False:
            print(f'Postprocessing not found for {tag}...')

        # print(f'url -> {url}')
        driver.get(url)
        # Press Alt + D to select browser address bar
        pyautogui.hotkey('alt', 'd')
        pyautogui.press('enter')

        time.sleep(2)

    # Close the WebDriver
    driver.quit()

    end = time.time()

print(f"Total runtime of the program is {end - begin}") 
