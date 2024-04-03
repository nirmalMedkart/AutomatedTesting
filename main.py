import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException


response = ['SUCCESS', 'ERROR']

def fileToDf(file_path):
    df = pd.read_excel(file_path)
    print(f"Excel file found and read.")
    return df


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


def storeOutput(main_url, res):
    data = {
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
    input_path = r"C:\Nirmal\Python\AutomatedTesting\input\input.xlsx"

    # Read the last entry's ID from output.xlsx
    output_df = pd.read_excel(r"C:\Nirmal\Python\AutomatedTesting\output\output.xlsx")

    # Check if the DataFrame is not empty
    if not output_df.empty:
        output_url = output_df.iloc[-1]['MAIN_URL']
    else:
        output_url = None


    df = fileToDf(input_path)
    
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)


    for index,row in df.iterrows():
        if (output_url is None or row['MAIN_URL'] == output_url):
            res = pageStatus(driver, row['MAIN_URL'])
            print(f"{res}: {row['MAIN_URL']}")
            storeOutput(row['MAIN_URL'], res)
            output_url = None
    driver.close()





