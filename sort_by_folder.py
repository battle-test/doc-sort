import os
import re
import shutil
import pandas as pd
from datetime import datetime

df = pd.read_excel("Задание.xlsx")
files = os.listdir(r'Документы')
dict_action = {"Документы\\Папка1": [], "Документы\\Папка2": [], "Документы\\Папка3": [], "Документы\\Папка3.1": []}

for folder in dict_action:
    # Create folders (Папка1, Папка2, Папка3)
    try:
        os.mkdir(f"{folder}")
    except FileExistsError:
        pass


def create_dict():
    # Create value where feature "Send" == 1
    for index in df.index:
        if df.iloc[:, 3][index] == 1:
            dict_action["Документы\\Папка1"].append((df.iloc[index, 0])[6:])
    # Create value where there are two consecutive numbers (e.g. 22, 33, 44)
    for code in df.iloc[:, 0]:
        if re.search(r"([1-9])\1$", code):
            dict_action["Документы\\Папка2"].append(int(code[6:]))
    # Range date 20.06.2020 - 10.07.2020
    start_data = datetime(2020, 6, 19)
    end_data = datetime(2020, 7, 11)
    # Create value range of dates
    for value in enumerate(df.iloc[:, 2]):
        index, data = value[0], value[1]
        try:
            data = datetime.strptime(data, r"%d.%m.%Y")
            if start_data < data < end_data:
                dict_action["Документы\\Папка3"].append((df.iloc[index, 0])[6:])
        # If the date is not valid, create a special folder and drop everything
        except ValueError:
            dict_action["Документы\\Папка3.1"].append((df.iloc[index, 0])[6:])

    for action in dict_action:
        dict_action[action] = list(map(int, dict_action[action]))


create_dict()

for file in files:
    # if COUNT FOLDERS > 0, was moved to one of the folders, also removed from the main folder.
    __COUNT_FOLDERS = 0
    # file before str:'КА_PA_000113_акт_13.07.2020.txt'
    # file after int:'113'
    try:
        index_file = int((file[6:].split('_'))[0][3:])
    except ValueError:
        pass


    def aggregation(directory):
        # directory = folder in dict_action (example:Документы\\Папка1)
        global __COUNT_FOLDERS
        for value in dict_action[f"{directory}"]:
            if value == index_file:
                # move in the folder(directory)
                shutil.copy2(f"Документы\\{file}", f"{directory}")
                __COUNT_FOLDERS += 1


    for folder in dict_action:
        aggregation(folder)

    if __COUNT_FOLDERS > 0:
        os.remove(f"Документы\\{file}")
