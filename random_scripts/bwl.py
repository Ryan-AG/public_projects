from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from pathlib import Path

my_path = Path(
    r'C:\Users\grane\Documents\Projects\Python\public_projects\random_scripts\geckodriver.exe')

driver = webdriver.Firefox(service=Service(my_path))
driver.get('http://www.google.com')
