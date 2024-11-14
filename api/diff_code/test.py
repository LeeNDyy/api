from quart import Quart, jsonify, request, send_file
import os
import aiofiles
import logging
import pandas as pd
import aiohttp
import asyncio

app = Quart(__name__)

# Настройка логирования
logging.basicConfig(level=logging.DEBUG)

# Лимит запросов
REQUEST_LIMIT = 100
REQUEST_DELAY = 60  # Задержка в секундах после 100 запросов

class ExcelHandler:
    def __init__(self, file_path):
        self.file_path = file_path
        self.dataframe = None

    def read_excel(self):
        """Чтение Excel файла и сохранение данных в dataframe."""
        try:
            self.dataframe = pd.read_excel(self.file_path)
            logging.info("Excel файл успешно прочитан.")
        except FileNotFoundError:
            raise Exception(f"Файл {self.file_path} не найден.")
        except Exception as e:
            raise Exception(f"Ошибка при чтении файла: {str(e)}")
    
    def add_coordinates_column(self, coordinates_column_name):
        """Добавление новой колонки с координатами."""
        if self.dataframe is not None:
            if coordinates_column_name not in self.dataframe.columns:
                self.dataframe[coordinates_column_name] = None
                logging.info(f"Колонка '{coordinates_column_name}' добавлена.")
            else:
                logging.info(f"Колонка '{coordinates_column_name}' уже существует.")
        else:
            raise Exception("Dataframe не загружен. Сначала вызовите метод read_excel.")
    
    async def save_excel(self):
        """Асинхронное сохранение Excel файла с изменениями в тот же файл."""
        try:
            self.dataframe.to_excel(self.file_path, index=False)
            logging.info(f"Файл успешно сохранен: {self.file_path}")
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

    async def get_coordinates(self, address):
        """Получение координат по адресу с использованием Yandex Geocoder API."""
        url = 'https://geocode-maps.yandex.ru/1.x/'
        params = {
            'geocode': address,
            'format': 'json',
            'apikey': self.api_key
        }

        async with aiohttp.ClientSession() as session:
            try:
                async with session.get(url, params=params) as response:
                    response.raise_for_status()  # Проверяем, что запрос успешен
                    data = await response.json()

                    if 'response' in data and data['response']['GeoObjectCollection']['featureMember']:
                        pos = data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos']
                        lon, lat = map(float, pos.split())  # Долгота, широта
                        return lat, lon
                    else:
                        logging.warning(f"Координаты не найдены для адреса: {address}")
                        return None, None
            except aiohttp.ClientError as e:
                logging.error(f"Ошибка при запросе к API для адреса {address}: {str(e)}")
                return None, None
            
@app.route('/upload', methods=['POST'])
async def upload_file():
    try:
        form = await request.files
        if 'file' not in form:
            return jsonify({'error': 'No file provided'}), 400
        
        file = form['file']
        file_path = os.path.join('./uploads', file.filename)
        
        # Чтение данных файла не нужно асинхронное, убираем await
        content = file.read()

        # Асинхронное сохранение файла
        async with aiofiles.open(file_path, 'wb') as f:
            await f.write(content)

        return jsonify({'message': 'File uploaded successfully', 'file_path': file_path})
    
    except Exception as e:
        logging.error(f"Ошибка загрузки файла: {e}")
        return jsonify({'error': str(e)}), 500


@app.route('/process', methods=['POST'])
async def process_addresses():
    try:
        data = await request.get_json()
        file_path = data.get('addresses')
        api_key_file = data.get('apikey')

        excel_handler = ExcelHandler(file_path)
        geocoder = AddressGeocoder(api_key_file)

        # Чтение Excel файла (синхронно)
        excel_handler.read_excel()
        excel_handler.add_coordinates_column('Координаты')

        request_count = 0  # Счётчик запросов

        for index, row in excel_handler.dataframe.iterrows():
            address = row.get('Адрес')
            if pd.isna(address):
                continue

            if pd.notna(excel_handler.dataframe.at[index, 'Координаты']):
                continue

            # Получение координат асинхронно
            latitude, longitude = await geocoder.get_coordinates(address)
            if latitude is not None and longitude is not None:
                excel_handler.dataframe.at[index, 'Координаты'] = f"{latitude}, {longitude}"

            request_count += 1

            # Проверка лимита запросов
            if request_count % REQUEST_LIMIT == 0:
                logging.info(f"Достигнут лимит запросов ({REQUEST_LIMIT}). Ожидание {REQUEST_DELAY} секунд...")
                await asyncio.sleep(REQUEST_DELAY)  # Ожидание перед отправкой следующей партии запросов

        # Асинхронное сохранение Excel файла
        await excel_handler.save_excel()
        logging.info(f"Обработка адресов завершена. Всего запросов: {request_count}")

        # Отправляем файл обратно пользователю
        return await send_file(file_path, as_attachment=True)
    
    except Exception as e:
        logging.error(f"Ошибка во время обработки адресов: {str(e)}")
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(port=5000)
