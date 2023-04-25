import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from tqdm import tqdm
from selenium.webdriver.common.by import By
import re
#import folium

#### Provide the absolute location of input file such as 'LATLNG1.xlsx' here ####
#### Kindly restrict the number of input data points to less than 100 on each run to prevent issues ####
data = pd.read_excel('/Users/hsomani/Documents/Python/LATLNG1.xlsx', keep_default_na=False)
output = data.copy()
output['Url'] = ['https://www.google.com/maps/search/' + i for i in output['Store'] ]
print('---------Search URL compiled from given data---------')

Url_With_Coordinates = []
#Store_Address = []

#### Web scraping in backgroup using Chrome and open given url ####
options = Options()
options = webdriver.ChromeOptions()
prefs = {'profile.default_content_setting_values': {'images':2, 'javascript':2}}
options.add_experimental_option('prefs', prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
driver.get("https://www.google.com")

print('---------Begin Google Map Scraping---------')
for url in tqdm(output.Url, leave=False):
   driver.get(url)
   Url_With_Coordinates.append(driver.find_element(By.CSS_SELECTOR, 'meta[itemprop=image]').get_attribute('content'))
   #Store_Address.append(driver.find_element(By.CSS_SELECTOR, 'meta[property=og:title]').get_attribute('content'))
    
driver.close()
print('---------URL with coordinates scraped---------')

#### Save meta URL of map location in separate excel file ####
output['Url_With_Coordinates'] = Url_With_Coordinates
substring = "size=200x200"
output.to_excel('url_output_1.xlsx')
print('---------URL Data is written to Excel File successfully---------')

print(output.head())

#### We will try to parse and get coordinates based on single or multiple stores from URL ####
for url in tqdm(output.Url_With_Coordinates, leave=False):
    index = output.index[output['Url_With_Coordinates'] == url]
    x = 0
    if substring in url:
        print("Multiple Stores")
        i = int(len(re.findall('(?=%7C)', url)))
        print(i)
        while x <= i:
            lat_col = 3+(x+1)*2
            #lat_col_name = "Store" + str(x+1) + "_lat"
            lng_col = 4+(x+1)*2
            #lng_col_name = "Store_" + str(x+1) + "_lng"
            #output.insert(lat_col, lat_col_name, "")
            output.loc[index, lat_col] = url.split('&markers=')[1].split('&sensor=')[0].split('%7C')[x].split('%2C')[0]
            #output.insert(lng_col, lng_col_name, "")
            output.loc[index, lng_col] = url.split('&markers=')[1].split('&sensor=')[0].split('%7C')[x].split('%2C')[1]
            x += 1

    else:
        print("Single Store")
        output.loc[index, 3] = url.split('?center=')[1].split('&zoom=')[0].split('%2C')[0]
        output.loc[index, 4] = url.split('?center=')[1].split('&zoom=')[0].split('%2C')[1]

#### Save output of coordinates in separate excel file ####
output.to_excel('latlng_output_1.xlsx')
print('---------Lat Long is written to Excel File successfully---------')
