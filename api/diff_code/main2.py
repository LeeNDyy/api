from flask import Flask, jsonify, request, send_file
import pandas as pd
import requests
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def allowed_file(filename):
    """Проверка допустимых расширений файлов."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


class ExcelHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        self.dataframe = None

    def read_excel(self):
        """Чтение Excel файла и сохранение данных в dataframe."""
        try:
            self.dataframe = pd.read_excel(self.file_path)
            print("Excel файл успешно прочитан.")
        except FileNotFoundError:
            raise Exception(f"Файл {self.file_path} не найден.")
        except Exception as e:
            raise Exception(f"Ошибка при чтении файла: {str(e)}")

    def add_coordinates_column(self, coordinates_column_name):
        """Добавление новой колонки с координатами."""
        if self.dataframe is not None:
            if coordinates_column_name not in self.dataframe.columns:
                self.dataframe[coordinates_column_name] = None
                print(f"Колонка '{coordinates_column_name}' добавлена.")
            else:
                print(f"Колонка '{coordinates_column_name}' уже существует.")
        else:
            raise Exception("Dataframe не загружен. Сначала вызовите метод read_excel.")

    def save_excel(self):
        """Сохранение Excel файла с изменениями в тот же файл."""
        try:
            self.dataframe.to_excel(self.file_path, index=False)
            print(f"Файл успешно сохранен: {self.file_path}")
        except Exception as e:
            raise Exception(f"Ошибка при сохранении Excel файла: {str(e)}")


class AddressGeocoder:
    def __init__(self, api_key_file):
        self.api_key = self.get_api_key(api_key_file)

    def get_api_key(self, file_path):
        """Чтение API ключа из файла."""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                api_key = file.read().strip()
            return api_key
        except FileNotFoundError:
            raise Exception(f"Файл {file_path} не найден.")
        except Exception as e:
            raise Exception(f"Ошибка чтения API ключа: {str(e)}")

    def get_coordinates(self, address):
        """Получение координат по адресу с использованием Yandex Geocoder API."""
        url = 'https://geocode-maps.yandex.ru/1.x/'
        params = {
            'geocode': address,
            'format': 'json',
            'apikey': self.api_key
        }

        try:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()

            if 'response' in data and data['response']['GeoObjectCollection']['featureMember']:
                pos = data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos']
                lon, lat = map(float, pos.split())
                return lat, lon
            else:
                print(f"Координаты не найдены для адреса: {address}")
                return None, None
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при запросе к API для адреса {address}: {str(e)}")
            return None, None


@app.route('/upload', methods=['POST'])
def upload_file():
    """Загрузка Excel файла."""
    if 'file' not in request.files:
        return jsonify({'error': 'Нет файла в запросе'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        return jsonify({'message': 'Файл загружен', 'file_path': file_path}), 200

    return jsonify({'error': 'Недопустимый формат файла. Разрешены только .xlsx файлы.'}), 400


@app.route('/process', methods=['POST'])
def process_addresses():
    """Обработка адресов и добавление координат."""
    data = request.get_json()
    file_path = data.get('file_path')
    api_key_file = 'apikey.txt'  # Укажем путь к файлу с API ключом
    max_requests_per_day = 900

    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'Неверный путь к файлу.'}), 400

    # Создаем экземпляры классов
    excel_handler = ExcelHandler(file_path)
    geocoder = AddressGeocoder(api_key_file)

    # Чтение Excel файла
    try:
        excel_handler.read_excel()
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    # Добавление столбца для координат
    coordinates_column_name = 'Координаты'
    try:
        excel_handler.add_coordinates_column(coordinates_column_name)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

    # Обрабатываем каждый адрес
    request_count = 0
    for index, row in excel_handler.dataframe.iterrows():
        address = row.get('Адрес')
        if pd.isna(address):
            continue

        if pd.notna(excel_handler.dataframe.at[index, coordinates_column_name]):
            continue

        if request_count >= max_requests_per_day:
            break

        latitude, longitude = geocoder.get_coordinates(address)
        if latitude is not None and longitude is not None:
            excel_handler.dataframe.at[index, coordinates_column_name] = f"{latitude}, {longitude}"
            request_count += 1

    # Сохранение файла
    try:
        excel_handler.save_excel()
        return jsonify({'message': 'Файл успешно обработан', 'file_path': file_path}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/download', methods=['GET'])
def download_file():
    """Скачивание обработанного файла."""
    file_path = request.args.get('file_path')
    if not file_path or not os.path.exists(file_path):
        return jsonify({'error': 'Файл не найден'}), 400

    return send_file(file_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
