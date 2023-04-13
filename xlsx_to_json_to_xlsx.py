from flask import Flask, request, Response
import base64
import json
import pandas as pd
from io import BytesIO
from werkzeug.serving import run_simple
import openpyxl


app = Flask(__name__)


@app.route('/upload', methods=['POST'])
def upload_file():
    # Получаем данные из POST запроса
    file_data = request.form['file']

    # Декодируем файл из base64 в бинарник
    file_binary_data = base64.b64decode(file_data)

    # считываем бинарник XLSX, приобразовываем в DataFrame и меняем NaN на ""
    df = pd.read_excel(file_binary_data).fillna('')

    # Преобразование DataFrame в словарь
    data = df.to_dict(orient='records')

    # Удаляем в словаре пробелы в начале и в конце каждого значения при условии, что значение - строка
    for record in data:
        for key, value in record.items():
            if isinstance(value, str):
                record[key] = value.strip()

                # ЭТОТ ФУНКЦИОНАЛ ПОКА НЕ НУЖЕН
                # находим все символы кроме рус.яз и знаков препинания и объединяем (нахождения артикула)
                # regex = r'[^а-яА-ЯёЁ\s\.,!?;:]+'
                # parts = re.findall(regex, record[key])
                # record[key] = ', '.join(parts)
    
    # Преобразование словаря в json
    json_data = json.dumps(data, ensure_ascii=False)

    # Возврат json
    return json_data



@app.route('/dowload', methods=['POST'])
def upload():
    # Получаем данные из POST запроса
    data = request.data

    # Преобразуем дату в json
    json_dict = json.loads(data)

    # принимаем список заголовок из списка headers
    headers = json_dict['headers']

    # Создаем DataFrame из JSON из списка data
    df = pd.DataFrame(json_dict['data'])

    # сортируем данные в обратном порядке (от новых к старым)
    df = df.sort_index(ascending=False)

    # Создаем буфер
    buffer = BytesIO()
    
    # Создаем файл excel
    writer = pd.ExcelWriter(buffer, engine='openpyxl')

    # Записываем данные на лист в excel и указываем заголовки headers
    df.to_excel(writer, sheet_name='Лист1', index=False, header=headers, na_rep='')

    # Закрываем объект записи
    writer.close()

    # Ставим указатель в начало буфера
    buffer.seek(0)

    # Ответ
    return Response(buffer, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': 'attachment;filename=modified_sample.xlsx'})


if __name__ == '__main__':
    run_simple('0.0.0.0', 5000, app, threaded=True)