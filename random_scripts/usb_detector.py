import os
from pathlib import Path
import pyautogui as gui
from time import time

clean_list = ['thiccboi']

while True:

    current_list = os.listdir('/Volumes')

    if current_list != clean_list:
        gui.alert(
            f'Warning, UBS: {current_list[0]} has been plugged in!')
        break
