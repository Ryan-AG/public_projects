from numpy import imag
import pyautogui as gui
from time import sleep
from pathlib import Path
import os
import basil
from random import choice
from glob import glob
import sqlite3

connection = sqlite3.connect('used_photos.db')

cursor = connection.cursor()
# p = r"D:\Basil\fale.jpg"
# fake_path = f"INSERT INTO used_photos VALUES ('{p}')"

# # cursor.execute('''
# #                CREATE TABLE used_photos
# #                (image path)
# #                '''
# #                )

# cursor.execute(fake_path)

connection.commit()

cursor.execute('''
               DELETE FROM used_photos;
               ''')

connection.commit()

# fetch = 'SELECT * FROM used_photos'


# l = [item for item in cursor.execute(fetch)]

# # l = list(cursor.execute('''
# #                SELECT * FROM used_photos
# #                '''))

# old_photo_list = [row[0] for row in cursor.execute(fetch)]
# test = 'asd'
# if test in old_photo_list:
#     print('Found it!')
# # for row in cursor.execute(fetch):
# #     print(row[0])

# print(old_photo_list)
# image_repository = r"D:\Basil"
# print(choice(image_list))

# while True:
#     sleep(1)
#     print(gui.position())
# basil.load_image()

# image_list = glob(f'{image_repository}\\*.*')
# images_list = os.listdir(image_repository)
# image_choice = choice(image_list)
# print(image_choice)
# image_choice = choice(image_repository)
# print(image_choice)
cursor.close()
connection.close()
