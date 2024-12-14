import pandas as pd
from PIL import Image


try:
    file_path = 'realisation_report.xlsx' # доступ для загрузки Excel фаайла
    # Какие колонки необходимо отобрать для дальнейших действий
    selected_columns = ['Товар', 'Кол-во', 'Реализовано на сумму, руб.', 'Доплата за счет Ozon, руб.']
    # определить какие заголовки и столбцы считывать
    datafile = pd.read_excel(file_path, sheet_name='Лист 1', header=12, usecols=selected_columns)
    print(datafile.head()) # проверка, вывод пяти первых строк
    # суммирование выручки и доплаты от озон, создание сводного файла
    pivot_table = pd.pivot_table(datafile, values=['Кол-во', 'Реализовано на сумму, руб.', 'Доплата за счет Ozon, руб.'],
                                 index=['Товар'], aggfunc='sum')
    print(pivot_table)
    pivot_table.to_excel('Pivot_report.xlsx') # выгрузка сводного файла
except FileNotFoundError:
    print('Файл не найден')


# image = Image.open('NY_ball.png', 'r')
# image.show()



