
from flask import Flask, jsonify, request
import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog
import time
import sys


sys.stdout.reconfigure(encoding='utf-8') #принудительная настрйока консоли 
def delete_simvol(text):
    return text.replace('\u200e', '') #удаление ненужных символов


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
            response = requests.get(url, params=params, timeout=10)  # Ограничение на ожидание ответа
            response.raise_for_status()  # Проверяем, что запрос успешен
            data = response.json()

            if 'response' in data and data['response']['GeoObjectCollection']['featureMember']:
                pos = data['response']['GeoObjectCollection']['featureMember'][0]['GeoObject']['Point']['pos']
                lon, lat = map(float, pos.split())  # Долгота, широта
                return lat, lon
            else:
                print(f"Координаты не найдены для адреса: {address}")
                return None, None
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при запросе к API для адреса {address}: {str(e)}")
            return None, None


def open_file_dialog():
    """Открытие диалогового окна для выбора файла."""
    root = tk.Tk()
    root.withdraw()  # Скрыть основное окно
    file_path = filedialog.askopenfilename(title="Выберите Excel файл", 
                                           filetypes=[("Excel файлы", "*.xlsx")])
    return file_path

def process_addresses(api_key_file, max_requests_per_day=300):
    # Открытие окна проводника для выбора Excel файла
    excel_file_path = open_file_dialog()
    if not excel_file_path:
        print("Файл не выбран.")
        return
    


    #Создаем экземпляры классов
    excel_handler = ExcelHandler(excel_file_path)
    geocoder = AddressGeocoder(api_key_file)

    #Чтение Excel файла
    try:
        excel_handler.read_excel()
    except Exception as e:
        print(e)
        return

    # Добавление столбца для координат, если его ещё нет
    coordinates_column_name = 'Координаты'
    try:
        excel_handler.add_coordinates_column(coordinates_column_name)
    except Exception as e:
        print(e)
        return

    # Считаем количество обработанных запросов
    request_count = 0

    # Обрабатываем каждый адрес, пропуская уже обработанные (где есть координаты)
    for index, row in excel_handler.dataframe.iterrows():
        address = row.get('Адрес')  # Имя столбца с адресами 
        if pd.isna(address):
            print(f"Строка {index}: пропущен, так как адрес пустой.")
            continue

        # Проверяем, если координаты уже есть, пропускаем
        if pd.notna(excel_handler.dataframe.at[index, coordinates_column_name]):
            print(f"Строка {index}: пропущен, так как координаты уже есть.")
            continue

        # Проверяем лимит запросов
        if request_count >= max_requests_per_day:
            print(f"Достигнут лимит в {max_requests_per_day} запросов за день. ")
            break

        # Получаем координаты для текущего адреса
        try:
            latitude, longitude = geocoder.get_coordinates(address)
            if latitude is not None and longitude is not None:
                excel_handler.dataframe.at[index, coordinates_column_name] = f"{latitude}, {longitude}"
                print(f"Обработан адрес: {address} -> {latitude}, {longitude}")
            else:
                print(f"Координаты не найдены для адреса: {address}")
            request_count += 1
        except Exception as e:
            print(f"Ошибка для адреса {address}: {str(e)}")

    # Сохранение Excel файла с обновлениями
    try:
        excel_handler.save_excel()
    except Exception as e:
        print(e)

def main():
    api_key_file = 'apikey.txt'  # Путь к файлу с API ключом

    # Запуск процесса с ограничением в 900 запросов в день
    process_addresses(api_key_file, max_requests_per_day=900)

if __name__ == '__main__':
    main()
