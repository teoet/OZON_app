import openpyxl
import sys
import os
import requests
import time
import re
from docx import Document
from docx.shared import RGBColor
import configparser
config = configparser.ConfigParser()

not_found_img = []
fetch_err = []
not_found_in_general = [] 

document = Document()

def check_internet_connection():
    url = 'https://www.freecodecamp.org/the-fastest-web-page-on-the-internet'
    timeout = 5
    retries = 10
    while retries > 0:
        try:
            requests.head(url, timeout=timeout)
            print("Internet connection available.", flush=True)
            return True
        except requests.exceptions.ConnectionError:
            print("No internet connection available. Retrying in 10 seconds...", flush=True)
            retries -= 1
            time.sleep(10)
        except requests.exceptions.ReadTimeout:
            print("Request timed out. Retrying in 10 seconds...", flush=True) 
            retries -= 1
            time.sleep(10)
    print("Unable to establish internet connection.", flush=True) 
    return False

def load_files():
    argc = len(sys.argv)
    if argc > 1:
        config.read(sys.argv[1])
    if argc > 2:
        output_file = sys.argv[2]
    if argc > 3:
        input_file1 = sys.argv[3]
    else:
        input_file1 = "../global/1.xlsx" 
    if argc > 4:
        input_file2 = sys.argv[4] 
    else:
        input_file2 = "../global/2.xlsx"
    if not os.path.isfile(input_file1):
        print("Файл {} не найден.".format(input_file1))
        return 0;
    if not os.path.isfile(sys.argv[1]):
        print("Файл {} не найден.".format(sys.argv[1]))
        return 0;
    if not os.path.isfile(input_file2):
        print("Файл {} не найден.".format(input_file2))
        return 0;
    if not os.path.isfile(output_file):
        print("Файл {} не найден.".format(output_file))
        return 0;
    print("Загрузка файлов...", flush=True)
    template_name = config.get('CATEGORY', 'template_name')
    category = config.get('CATEGORY', 'category')
    input_workbook1 = openpyxl.load_workbook(input_file1)
    input_workbook2 = openpyxl.load_workbook(input_file2)
    output_workbook = openpyxl.load_workbook(output_file)
    input_sheet1 = input_workbook1.active
    input_sheet2 = input_workbook2.active
    output_sheet = output_workbook.worksheets[4]
    first_load(input_sheet1, input_sheet2, output_sheet, output_workbook, output_file, category, template_name)

def log(message, color=None):
    RED = RGBColor(255, 0, 0)
    GREEN = RGBColor(0, 255, 0)
    YELLOW = RGBColor(230, 115, 0)
    BLACK = RGBColor(0, 0, 0)
    if color == 'red':
        paragraph = document.add_paragraph('- ', style='Normal')
        run = paragraph.add_run(message)
        run.font.color.rgb = RED
    elif color == 'yellow':
        paragraph = document.add_paragraph('- ', style='Normal')
        run = paragraph.add_run(message)
        run.font.color.rgb = YELLOW
    elif color == 'green':
        paragraph = document.add_paragraph('- ', style='Normal')
        run = paragraph.add_run(message)
        run.font.color.rgb = GREEN
    else:
        paragraph = document.add_paragraph('- ', style='Normal')
        run = paragraph.add_run(message)
        run.font.color.rgb = BLACK


def get_image_url(index):
    result_list = []
    url = f'https://maxmaster.ru/index.php?dispatch=products.cp_live_search&q={index}&search_input_id=&result_ids=live_reload_box&is_ajax=1'
    respone = None
    try:
        response = requests.get(url,timeout=10)
    except Exception as e:
        fetch_err.append(index)
        log(f"\nЧто то пошло не так при поиске фотографии товара {index}", 'red')
        if check_internet_connection() is True: return None
    if response.status_code == 200:
        result = response.json()
        start = result['html']['live_reload_box'].find('a href')
        end = result['html']['live_reload_box'].find(">", result['html']['live_reload_box'].find('a href'))
        card_url = result['html']['live_reload_box'][(start + 8):(end - 2)]
        try:
            response = requests.get(card_url, timeout=10)
        except Exception as e:
            e = str(e)
            if "Read timed out" in e or "MaxRetryError" in e or check_internet_connection is not True:
                fetch_err.append(index)
                return None
            elif "Нет элементов для отображения" in e.args[0]:
                log(f"\nТовар с данным артикулом {index} не найден", 'red')
                not_found_img.append(index)
            else:
                fetch_err.append(index)
                return None
    else:
        fetch_err.append(index)
        log(f"\nЧто то пошло не так при поиске фотографии товара {index}", 'yellow')
        return None
    if response.status_code == 200:
        result = response.text
        pattern = r'ty-pict\s+cm-image'
        matches = re.finditer(pattern, result)
        for count, match in enumerate(matches):
            start_index = match.start()
            end_index = match.end()
            img_url_start = result.find("src=", start_index)
            img_url_start += 5
            img_url_end = result.find('"', (img_url_start))
            img_url = result[img_url_start:img_url_end]
            if '661/661' in img_url:
                result_list.append({'count': count, 'img_url': img_url})
    else:
        fetch_err.append(index)
        log(f"\nЧто то пошло не так при поиске фотографии товара {index}", 'yellow')
        return None
    return result_list


def calculate_price(articul, price_value):
    if price_value is None:
        log(f"\nЗначение \"Цена\" у наименования {articul} пустое, пожалуйста заполните вручную. Установлено стандартное значение 0.", 'red')
        return 0
    elif price_value < 1000:
        price_value += (60 * price_value) / 100
    elif price_value >= 1000 and price_value < 10000:
        price_value += (57 * price_value) / 100
    else:
        price_value += (50 * price_value) / 100
    return round(price_value)


def calculate_weight(articul, weight_value):
    if weight_value is None:
        log(f"\nЗначение \"Вес\" у наименования {articul} пустое, пожалуйста заполните вручную. Установлено стандартное значение 0.", 'red')
        return 0
    if isinstance(weight_value, str):
        weight_value = int(weight_value)
    elif weight_value <= 0.05:
        weight_value += (100 * weight_value) / 100
    elif weight_value > 0.05 and weight_value <= 0.1:
        weight_value += (50 * weight_value) / 100
    elif weight_value > 0.05 and weight_value <= 0.15:
        weight_value += (30 * weight_value) / 100
    elif weight_value > 0.15 and weight_value <= 0.2:
        weight_value += (20 * weight_value) / 100
    elif weight_value > 0.2 and weight_value <= 0.5:
        weight_value += (15 * weight_value) / 100
    elif weight_value > 0.5 and weight_value <= 0.9:
        weight_value += (10 * weight_value) / 100
    elif weight_value > 0.9:
        weight_value += (5 * weight_value) / 100
    return round(weight_value * 1000)


def check_size(articul, size):
    if size is None:
        log(f"Значение \"Размер\" у наименования {articul} пустое, пожалуйста заполните вручную. Установлено стандартное значение 0.", 'green')
        return 0
    size += (10 * size) / 100
    return round(size * 10)




def second_load(input_sheet1, input_sheet2, output_sheet, output_workbook, output_file, template_name):
    for row in range(4, output_sheet.max_row + 1):
        row_num = row
        found = False;
        needed_articul = output_sheet.cell(row=row, column=2).value
        if(needed_articul is None): continue
        for row in range(1, input_sheet2.max_row + 1):
            cell_value = input_sheet2.cell(row=row, column=1).value
            if(cell_value is None): continue
            if needed_articul == cell_value:
                found = True;
                weight_value = calculate_weight(needed_articul, input_sheet2.cell(row=row, column=16).value)
                length_value = check_size(needed_articul, input_sheet2.cell(row=row, column=17).value)
                height_value = check_size(needed_articul, input_sheet2.cell(row=row, column=18).value)
                width_value = check_size(needed_articul, input_sheet2.cell(row=row, column=19).value)
                brand_value = str(input_sheet2.cell(row=row, column=2).value)
                if brand_value is None:
                    log(f"Значение \"Торговая марка\" у наименования {needed_articul} пустое, пожалуйста заполните вручную.", 'red')
                made_in_value = str(input_sheet2.cell(row=row, column=28).value)
                if made_in_value is None:
                    log(f"Значение \"Страна Изготовителя\" у наименования {needed_articul} пустое, пожалуйста заполните вручную.")
                key_words_value = input_sheet2.cell(row=row, column=8).value
                if key_words_value is None:
                    log(f"Значение \"Ключевые слова\" у наименования {needed_articul} пустое, пожалуйста заполните вручную.")
                bar_code_value = input_sheet2.cell(row=row, column=29).value
                if bar_code_value is None:
                    log(f"Значение \"Шрих-Код\" у наименования {needed_articul} пустое, пожалуйста заполните вручную.")
                annotation_value = " " if input_sheet2.cell(row=row, column=12).value is None else input_sheet2.cell(row=row, column=12).value
                annotation_value += " " if input_sheet2.cell(row=row, column=13).value is None else input_sheet2.cell(row=row, column=13).value
                annotation_value += " " if input_sheet2.cell(row=row, column=11).value is None else input_sheet2.cell(row=row, column=11).value
               
                weight_tmp = int(config.get("VALUES", "weight")) 
                length_tmp = int(config.get("VALUES", "length")) 
                width_tmp = int(config.get("VALUES", "width")) 
                brand_tmp = int(config.get("VALUES", "brand")) 
                made_in_tmp = int(config.get("VALUES", "made_in"))
                height_tmp = int(config.get("VALUES", "height")) 
                model_tmp = int(config.get("VALUES", "model")) 
                complect_tmp = int(config.get("VALUES", "complect")) 
                warranty_tmp= int(config.get("VALUES", "warranty")) 
                key_tmp = int(config.get("VALUES", "key_words")) 
                bar_tmp = int(config.get("VALUES", "bar_code")) 
                annot_tmp = int(config.get("VALUES", "annotation")) 
                
                if weight_tmp != 0:
                    output_sheet.cell(row=row_num, column=weight_tmp).value = weight_value
                if length_tmp != 0:
                    output_sheet.cell(row=row_num, column=length_tmp).value = length_value
                if height_tmp != 0:
                    output_sheet.cell(row=row_num, column=height_tmp).value = height_value 
                if width_tmp != 0:
                    output_sheet.cell(row=row_num, column=width_tmp).value = width_value 
                if brand_tmp != 0:
                    output_sheet.cell(row=row_num, column=brand_tmp).value = brand_value 
                if made_in_tmp != 0:
                    output_sheet.cell(row=row_num, column=made_in_tmp).value = made_in_value
                if complect_tmp != 0:
                    output_sheet.cell(row=row_num, column=complect_tmp).value = complect_value 
                if warranty_tmp != 0:
                    output_sheet.cell(row=row_num, column=warranty_tmp).value = warranty_value
                if key_tmp != 0:
                    output_sheet.cell(row=row_num, column=key_tmp).value = key_words_value
                if bar_tmp != 0:
                    output_sheet.cell(row=row_num, column=bar_tmp).value = bar_code_value
                if annot_tmp != 0:
                    output_sheet.cell(row=row_num, column=annot_tmp).value = annotation_value

                break
        if found is False:
            not_found_in_general.append(needed_articul)
            log(f"Товар с данным артикулом {needed_articul} не найден в полной выгрузке. Значения не добавлены.", 'red')
        progress = (row_num / output_sheet.max_row) * 100
        if progress % 3 == 0:
            print(f"Перенос значений из второго файла: {progress:.2f}%", flush=True)
    print("\nПеренос значений из второго файла завершен.", flush=True)
    output_workbook.save(output_file)
    log("Не найденные фотографии: " + str(not_found_img))
    log("Ошибка при поиске фотографии: " + str(fetch_err), 'yellow')
    log("Не найдено в полной выгрузке: " + str(not_found_in_general), 'red')
    document.save(f'{template_name}.docx')

def first_load(input_sheet1, input_sheet2, output_sheet, output_workbook, output_file, category, template_name):
    row_num = 4
    max_row = input_sheet1.max_row
    for row in range(1, max_row + 1):
        if (input_sheet1.cell(row=row, column=16).value == category):
            articul_value = input_sheet1.cell(row=row, column=1).value
            name_value = input_sheet1.cell(row=row, column=2).value
            if name_value is None:
                log(f"Значение \"Название\" у наименования {articul_value} пустое, пожалуйста заполните вручную.", 'yellow')
            brand_value = str(input_sheet1.cell(row=row, column=9).value).upper()
            if brand_value is None:
                log(f"Значение \"Торговая Марка\" у наименования {articul_value} пустое, пожалуйста заполните вручную.", 'yellow')
            price_value = calculate_price(articul_value, input_sheet1.cell(row=row, column=4).value)
            model_name = input_sheet1.cell(row=row, column=18).value
            if model_name is None:
                model_name = articul_value
            type_value = input_sheet1.cell(row=row, column=17).value
            if type_value is None:
                log(f"Обязательное значение \"Тип товара\" у наименования {articul_value} пустое, пожалуйста заполните вручную.", 'red')
            img_url = get_image_url(articul_value)
            if img_url is not None:
                main_img = [item['img_url'] for item in img_url if item['count'] == 0]
                output_sheet.cell(row=row_num, column=14).value = main_img[0]
                other_imgs = [item['img_url'] for item in img_url if item['count'] != 0]
                output_sheet.cell(row=row_num, column=15).value = ', '.join(other_imgs)
            

            articul_tmp = int(config.get("VALUES", "articul")) 
            name_tmp = int(config.get("VALUES", "name")) 
            price_tmp = int(config.get("VALUES", "price"))
            price2_tmp = int(config.get("VALUES", "price2")) 
            nds_tmp = int(config.get("VALUES", "nds")) 
            model_tmp = int(config.get("VALUES", "model")) 
            type_tmp = int(config.get("VALUES", "type")) 
            template_tmp = int(config.get("VALUES", "com_type")) 


            if articul_tmp != 0:
                output_sheet.cell(row=row_num, column=articul_tmp).value = articul_value
            if name_tmp != 0:
                output_sheet.cell(row=row_num, column=name_tmp).value = name_value
            if price_tmp != 0:
                output_sheet.cell(row=row_num, column=price_tmp).value = price_value
            if price2_tmp != 0:
                output_sheet.cell(row=row_num, column=price2_tmp).value = price_value 
            if nds_tmp != 0:
                output_sheet.cell(row=row_num, column=nds_tmp).value = "НДС не облагается"
            if template_tmp != 0:
                output_sheet.cell(row=row_num, column=template_tmp).value = template_name 
            if model_tmp != 0:
                output_sheet.cell(row=row_num, column=model_tmp).value = model_name
            if type_tmp != 0:
                output_sheet.cell(row=row_num, column=type_tmp).value = type_value

            row_num += 1
        progress = (row / input_sheet1.max_row) * 100
        print(f"Перенос значений из первого файла: {progress:.2f}%", flush=True)
    print("\nПеренос значений из первого файла завершен.", flush=True)
    output_workbook.save(output_file)
    second_load(input_sheet1, input_sheet2, output_sheet, output_workbook, output_file, template_name)

load_files()

