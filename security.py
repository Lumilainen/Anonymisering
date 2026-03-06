import os
import shutil

TEMP_FOLDER = "temp"


def ensure_temp():

    if not os.path.exists(TEMP_FOLDER):
        os.makedirs(TEMP_FOLDER)


def clean_temp():

    if os.path.exists(TEMP_FOLDER):

        for file in os.listdir(TEMP_FOLDER):

            path = os.path.join(TEMP_FOLDER, file)

            try:
                os.remove(path)
            except:
                pass