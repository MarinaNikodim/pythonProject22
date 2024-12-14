import pandas as pd
from PIL import Image
from PIL import ImageDraw, ImageFont


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


with Image.open('NY_ball.png', 'r') as image:
    #image.show()
    print(image.format)
    print(image.size)
    print(image.mode)
    box = (75, 230, 1090, 1260)
    cropped_image = image.crop(box)
    cropped_image.save('NYformat_ball.png')
    image_resize = image.resize((800, 900))
    image_resize.save('NY_resided_ball.png')

with Image.open('NY_tree.png', 'r') as image2:
    #image2.show()
    bandw = image2.convert('L') # преобразован в черно-белую картинку
    bandw.save('NY_tree_bw.jpg') # изменен формат с png на jpg
    draw = ImageDraw.Draw(image2)
    font = ImageFont.load_default(size=50)
    draw.text((70, 70), 'Happy New Year!!!', font=font, fill='purple')
    image2.save('greeting_card.png')


def create_collage(images, collage_width):# создание коллажа из фотографий
    collage_height = max(image.size[1] for image in images)
    collage = Image.new('RGB', (collage_width, collage_height), 'white')
    x_offset = 0
    for img in images:
        collage.paste(img, (x_offset, 0))
        x_offset += img.size[0]
    return collage


image_path = ['NY_tree.png', 'NYformat_ball.png']
images = [Image.open(img_path) for img_path in image_path]
collage = create_collage(images, 1900)
collage.save('collage.png')
collage.show()












