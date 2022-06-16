import pyautogui as gui
from time import sleep
import os
from pathlib import Path
import wmi
from random import choice
from glob import glob
import sqlite3

connection = sqlite3.connect('used_photos.db')
cursor = connection.cursor()

""" ONLY EXECUTED ONCE
cursor.execute('''
               CREATE TABLE used_photos
               (image path)
               '''
               )
"""

image_repository = Path(r"D:\Basil")
image_list = glob(f'{image_repository}\\*.*')

slack_path = Path(
    r'C:\Users\grane\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Slack Technologies Inc\Slack.lnk'
)

c = wmi.WMI()  # opens wmi connection


def open_slack(file_path: Path):
    """Attempts to open Slack.

    Args:
        file_path (Path): Path to the slack.exe file.
    """
    try:
        os.startfile(slack_path)
    except Exception as e:
        print(e)


def select_channel():
    """Moves curser to the side panel, scrolls to the top, and selects basil-and-friends."""
    gui.moveTo(104, 421)
    sleep(1)
    gui.scroll(1000)
    sleep(1)
    gui.click()
    sleep(1)


def terminate_slack():
    for process in c.Win32_Process():
        try:
            if process.name == 'slack.exe':
                process.Terminate()
        except Exception:
            print('Closed')


def check_db(image_choice):

    db_list = []
    query = f'''
    SELECT * FROM used_photos
    '''
    for row in cursor.execute(query):
        db_list.append(row[0])

    return image_choice in db_list


def update_db(image_choice):
    query = f'''
    INSERT INTO used_photos VALUES ('{image_choice}')
    '''
    cursor.execute(query)
    connection.commit()


def main():
    open_slack(slack_path)

    sleep(3)

    select_channel()

    searching = True
    terminate_counter = 0

    while searching:
        image_choice = choice(image_list)
        if not check_db(image_choice) or terminate_counter == 100:
            searching = False
        else:
            terminate_counter += 1

    with gui.hold('CTRL'):
        gui.press('u')
    gui.typewrite(image_choice)
    gui.press('enter')

    sleep(3)

    # gui.press('enter')

    sleep(0.5)

    terminate_slack()

    update_db(image_choice)

    cursor.close()
    connection.close()


if __name__ == '__main__':
    main()
