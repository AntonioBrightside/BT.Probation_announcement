import pandas as pd
from tkinter import *
from datetime import date

# Display settings
pd.options.display.width = None
pd.options.display.max_columns = None


# Preprocessing
def clear(df):
    df = df.dropna(subset='Name')
    df = df.fillna('')
    df = df.drop(index=df[df['№'] == '№'].index)

    df['Start date'] = pd.to_datetime(df['Start date'], format='%Y-%m-%d').dt.date
    df['Probation period'] = pd.to_datetime(df['Probation period'], format='%Y-%m-%d').dt.date
    return df


# Calculating
def calc(df):
    df['days_left'] = df['Probation period'] - date.today()
    return df


# Probation period is ending
def ppEnding(df):
    first = pd.to_timedelta(7, unit='D')
    end = pd.to_timedelta(0, unit='D')
    counter = 0
    dct = {}
    for meatbag in df['days_left']:
        if end < meatbag <= first:
            dct_part = {'name': df.iloc[counter]["Name"],
                        'dl': df.iloc[counter]["days_left"].days,
                        'pp': df.iloc[counter]["Probation period"]}
            dct_part = {counter: dct_part}
            dct.update(dct_part)
        counter += 1
    return dct


# Text to insert
def text(**kwargs):
    for key in kwargs.keys():
        text = f'{kwargs[key]["name"]} - '


# Start
if __name__ == '__main__':

    # Loading file data
    try:
        df = pd.read_excel(r"Y:\ИС проба.xlsx", header=1)  # мой файл: C:\Users\Antonio\Desktop\ИС проба.xlsx
    except FileNotFoundError:
        print('Неправильно указан адрес файла / имя файла. Обратитесь к разработчику для замены адреса')

    # File reading
    df = clear(df)
    df = calc(df)
    dictionary = ppEnding(df)

    # Interface
    root = Tk()

    L = Label(root, text='Сотрудники, у которых заканчивается испытатательный срок в течение недели: ')
    L.config(font=10)
    L.grid(row=0, column=0, columnspan=2, sticky=W)
    T = Text(root, width=200, height=100, wrap=WORD, font=15)

    try:
        for key in dictionary.keys():
            if dictionary[key]['dl'] in [2, 3, 4]:
                T.insert(0.0, f'\n {dictionary[key]["name"]} - {dictionary[key]["dl"]} дня ({dictionary[key]["pp"]})')
            elif dictionary[key]['dl'] == 1:
                T.insert(0.0, f'\n {dictionary[key]["name"]} - {dictionary[key]["dl"]} день ({dictionary[key]["pp"]})')
            else:
                T.insert(0.0, f'\n {dictionary[key]["name"]} - {dictionary[key]["dl"]} дней ({dictionary[key]["pp"]})')

            length = 20 + 25 + len(dictionary.keys())*20

        T.grid(row=1, column=0, columnspan=10, sticky=W)

        root.title('')
        root.geometry(f'600x{length}')
        root.mainloop()

    except NameError:
        exit()


