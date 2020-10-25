import pandas as pd
import os
from datetime import datetime
start_program = datetime.now()

try:
    os.mkdir("Документы")
except FileExistsError:
    pass

df = pd.read_excel("Задание.xlsx")
code_KA, docs, date = df.iloc[:, 0], df.iloc[:, 1], df.iloc[:, 2]

for index in df.index:
    docs[index] = docs[index].split(',')
    for i in range(len(docs[index])):
        file = open(f"Документы/КА_{code_KA[index]}_{docs[index][i]}_{date[index]}.txt", "w")

end_program = datetime.now()
time_program = end_program - start_program
print (time_program)