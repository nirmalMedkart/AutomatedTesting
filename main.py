import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException


response = ['SUCCESS', 'ERROR']

def fileToDf(folder_path):
    files = os.listdir(folder_path)

    # Filter out Excel and CSV files
    excel_files = [file for file in files if file.endswith('.xlsx') or file.endswith('.xls')]
    csv_files = [file for file in files if file.endswith('.csv')]

    # Read Excel file if present
    if len(excel_files) == 1:
        file_name = excel_files[0]
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_excel(file_path)
        print(f"Excel file '{file_name}' found and read.")
        return df

    # Read CSV file if present and no Excel files found
    elif len(csv_files) == 1:
        file_name = csv_files[0]
        file_path = os.path.join(folder_path, file_name)
        df = pd.read_csv(file_path)
        print(f"CSV file '{file_name}' found and read.")
        return df

    else:
        print("No Excel or CSV files found in the folder.")
        return None


def pageStatus(driver, main_url):
    xPath = '//*[@id="CardID"]/div/div[2]/div/div[4]/div[1]/h1'
    try:
        driver.get(main_url)
        time.sleep(2)
        data = driver.find_element(By.XPATH,xPath).text

        if data is not None:
            return response[0]
        else:
            return response[1]
    except NoSuchElementException:
        return response[1]
    except StaleElementReferenceException:
        return response[1]


def storeOutput(path, id, slug, cms_url, main_url, res):
    data = {
        'ID': [id],
        'SLUG': [slug],
        'CMS_URL': [cms_url],
        'MAIN_URL': [main_url],
        'RESPONSE': [res]
    }
    df1 = pd.DataFrame(data).reset_index(drop = True)

    df = pd.read_excel(r'C:\Nirmal\Python\AutomatedTesting\output\output.xlsx').reset_index(drop = True)
    df = pd.concat([df, df1]).reset_index(drop = True)
    df.to_excel(r'C:\Nirmal\Python\AutomatedTesting\output\output.xlsx', index = False)


# Main method
if __name__ == "__main__":

    # set your input and output folder path
    input_path = f"C:/Nirmal/Python/AutomatedTesting/input/"
    outputPath = f"C:/Nirmal/Python/AutomatedTesting/output/"

    # Read the last entry's ID from output.xlsx
    output_df = pd.read_excel(r"C:\Nirmal\Python\AutomatedTesting\output\output.xlsx")

    # Check if the DataFrame is not empty
    if not output_df.empty:
        output_id = output_df.iloc[-1]['ID']
    else:
        output_id = None


    df = fileToDf(input_path)
    
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)


    for index,row in df.iterrows():
        if (output_id is None or row['ID'] == output_id):
            res = pageStatus(driver, row['MAIN_URL'])
            print(f"{res}: {row['MAIN_URL']}")
            storeOutput(outputPath,row['ID'], row['SLUG'], row['CMS URL'], row['MAIN_URL'], res)
            output_id = None
    driver.close()





