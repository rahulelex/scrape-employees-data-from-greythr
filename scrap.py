import requests
import json
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import configparser
import os
import calendar
import pandas as pd
import matplotlib.pyplot as plt
import configparser


class DataExporter:
    def __init__(self):
        # Create a ConfigParser object
        self.config = configparser.ConfigParser()
        # Read the configuration file
        self.config.read('config.ini')

        # Access the configuration values
        self.xlsx_file_path = self.config.get('DEFAULT', 'xlsx_file_name')
        self.username = self.config.get('database', 'username')
        self.password = self.config.get('database', 'password')
        self.driver = None
        self.access_token = None
        self._ga = None
        self._gid = None
        self.max_number_of_pages= 10
        self.login_url= self.config.get('DEFAULT', 'login_url')
        self.employee_url= self.config.get('DEFAULT', 'employee_url')
        self.data_url= self.config.get('DEFAULT', 'data_url')
        self.page_size= self.config.get('DEFAULT', 'page_size')


    def login(self):
        # Configure Chrome options for headless browsing
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # Run Chrome in headless mode
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chromedriver_path = 'chromedriver_linux64/chromedriver'
        service = Service(chromedriver_path)
        driver = webdriver.Chrome(service=service, options= chrome_options)
        print('Login process initiated')
        driver.get(self.login_url)
        time.sleep(5)
        
        username_input = driver.find_element(By.ID, "username")
        password_input = driver.find_element(By.ID,'password')
        username_input.send_keys(self.username)
        password_input.send_keys(self.password)

        button = driver.find_element("css selector", 'button[type="submit"].bg-primary')
        button.click()

        time.sleep(2)

        # Navigate to the URL
        driver.get(self.employee_url)

        cookies = driver.get_cookies()

        for x in cookies:
            name= x.get('name')
            if name == 'access_token':
                self.access_token= x.get('value')
            elif name == '_ga':
                self._ga= x.get('value')
            elif name == '_gid':
                self._gid= x.get('value')

        time.sleep(2)

        if self._ga and self._gid and self.access_token is not None:
            print("LOGIN SUCCESS")
        else:
            print("LOGIN FAILED")
            exit

    def fetch_data(self):
        payload = json.dumps({
        "cat::search": {}
        })

        headers = {
        'content-type': 'application/json',
        'cookie': 'access_token='+self.access_token+'; _ga='+self._ga+'; _gid='+self._gid+'; _dc_gtm_UA-642192-18=1; _hjIncludedInSessionSample=1',
        }

        session= requests.session()

        if os.path.exists(self.xlsx_file_path):
            print('File already exists, removing it.')
            os.remove(self.xlsx_file_path)

        # Create an empty DataFrame
        df = pd.DataFrame()
        dfs= []

        for index in range(self.max_number_of_pages):
            response = session.request("POST", self.data_url+str(index)+'&pageSize='+str(self.page_size), headers=headers, data=payload)
            y = json.loads(response.text)

            if y.get('results'):
                for z in y.get('results'):
                    e_data = {
                        'Name': z.get('name'),
                        'DOB': z.get('dob'),
                        'Designation': z.get('c_designation'),
                        'Employee No': z.get('employeeno'),
                        'Employee ID': z.get('employeeid'),
                        'Email': z.get('email')
                    }
                    
                    # Create a DataFrame from the current iteration data
                    df = pd.DataFrame([e_data])
                    
                    # Append the DataFrame to the list
                    dfs.append(df)

        session.close()

        # Concatenate all DataFrames in the list
        df_merged = pd.concat(dfs, ignore_index=True)

        # Write the merged DataFrame to an Excel file
        df_merged.to_excel(self.xlsx_file_path, index=False)

        print("Data appended to the Excel file successfully.")

        # Read the XLSX file to remove duplicate data
        df = pd.read_excel(self.xlsx_file_path, engine='openpyxl')

        # Remove duplicate rows based on all columns
        df.drop_duplicates(inplace=True)

        # Write the updated DataFrame back to the Excel file
        df.to_excel(self.xlsx_file_path, index=False)

        print('Data cleaned - removed duplicate data')

        
    def show_charts(self):
        df = pd.read_excel(self.xlsx_file_path, engine='openpyxl')

        # Fetch the data from the specified column
        column_data = df['Designation']

        # Count the occurrences of each value
        value_counts = column_data.value_counts()

        # Display bar graph for column values
        value_counts.plot(kind='bar')
        fig1 = plt.figure(1)
        plt.title('Designation chart')
        plt.xlabel('Values')
        plt.ylabel('Count')

        # Pie chart for month of employees born in
        # Update column with month
        df['Month'] = pd.to_datetime(df['DOB'], format='%d %b').dt.month

        df['Month'] = df['Month'].replace('', float('nan')).astype(float).astype('Int64')

        # Replace numeric month values with month names
        df['Month'] = df['Month'].apply(lambda x: calendar.month_name[x] if pd.notnull(x) else '')

        # Count the occurrences of each month
        month_counts = df['Month'].value_counts()

        # Plot the pie chart
        fig2 = plt.figure(2)
        plt.pie(month_counts, labels=month_counts.index, autopct='%1.1f%%')
        plt.title('Month-wise Birthday discribution')

        # Show both charts
        plt.show()


def main():
    exporter = DataExporter()
    exporter.login()
    exporter.fetch_data()
    exporter.show_charts()

if __name__ == '__main__':
    main()
