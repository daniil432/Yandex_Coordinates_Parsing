import datetime
import os
import random
import time
import bs4
import re
import openpyxl
import pandas as pd
from requests_html import HTMLSession
from pandas.core.common import flatten
import numpy as np


def read_dataset(path):
    # Считываем данные из файла
    # Берём два столбца с широтой и долготой, пропуская пустые строки и строки с наименованиями
    coordinates = pd.read_excel(path, "ГПН-Б_сравнение", header=0, usecols='I:J', skiprows=1)
    coordinates = coordinates.loc[0:]
    # Считываем координаты баз, пропуская лишние строки и столбцы
    bases = pd.read_excel(path, "КМ", nrows=2, usecols='F:L')
    # Считываем предполагаемые названия объектов на изучаемых координатах
    names = pd.read_excel(path, "ГПН-Б_сравнение", header=0, usecols='D', skiprows=1)
    names = names.loc[0:]
    # Переводим полученные таблицы данных в списки
    coordinates = coordinates.values.tolist()
    bases = bases.values.tolist()
    names = names.values.tolist()
    print(len(coordinates), coordinates,)
    print(len(bases), bases,)
    print(len(names), names,)
    if len(coordinates) == len(names):
        print('Длины совпадают, процесс должен идти нормально')
    else:
        print('Число точек не совпало с числом наименований, проверьте код выше со считыванием данных либо сам файл')
    return coordinates, bases, names


def parse_km_m(s):
    # Извлекаем из посылаемой сюда строки длины маршрутов в километрах или метрах
    # Получаем из html текст вида км или м
    route_str = re.findall(r'[А-я]+', s)
    # Если маршрут в метрах, переводим в километры
    if 'м' in route_str:
        # Ищем число метров, которое может быть как целым, так и нецелым. Ищем в html-элементах через регулярные выражения
        route_str_float = re.findall(r'>(\d+\,\d+) м', s)
        route_str = re.findall(r'>(\d+) м', s)
        # Если есть дробные значения, то подменяем запятую в них на точку, переводим из str в float и делим на 1000
        if route_str_float != []:
            for i in range(len(route_str_float)):
                route_str_float[i] = route_str_float[i].replace(',', '.')
                route_str_float[i] = float(route_str_float[i])
                route_str_float[i] = route_str_float[i] / 1000
        # Если есть целые значения, переводим их в км
        if route_str != []:
            for i in range(len(route_str)):
                route_str[i] = float(route_str[i])
                route_str[i] = route_str[i] / 1000
            # Получаем список целых и нецелых величин маршрута
        route_str = route_str + route_str_float
    # Поиск длины пути, если путь не в метрах, а в километрах
    # Ищем число метров, которое может быть как целым, так и нецелым. Ищем в html-элементах через регулярные выражения
    elif 'км' in route_str and 'м' not in route_str:
        route_str_float = re.findall(r'>(\d+\,\d+) км', s)
        route_str = re.findall(r'>(\d+) км', s)
        # Если есть целые значения, переводим их из str в float
        if route_str != []:
            for i in range(len(route_str)):
                route_str[i] = float(route_str[i])
        # Если есть дробные значения, то подменяем запятую в них на точку, переводим из str в float и делим на 1000
        if route_str_float != []:
            for j in range(len(route_str_float)):
                route_str_float[j] = route_str_float[j].replace(',', '.')
                route_str_float[j] = float(route_str_float[j])
        # Получаем список целых и нецелых величин маршрута
        route_str = route_str + route_str_float
    else:
        print('??? Не метры и киломерты ???')
        route_str = []
    print(route_str)
    return route_str


def parse_for_routes(coordinates, bases, mode):
    # Поиск величины пути маршрута от одной нефтебаз до АЗС
    # Заголовки, с которыми точно работает код. Нужны при отправке запросов яндексу
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive',
        'Host': 'market.yandex.ru',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36',
    }
    # Постоянная часть ссылки
    const = "https://yandex.ru/maps/"
    # Список станет матрицей, содержащей все маршруты от каждой нефтебазы до каждой АЗС
    all_routes = []
    # Такая же матрица, но со ссылками на маршруты
    link_matrix = []
    # Счётчик шагов
    step = 0
    # Проходимся по координатам АЗС
    for azs in range(len(coordinates)):
        all_bases_to_azs = []
        link_temp = []
        # Проходимся по базам
        for coord in range(len(bases[0])):
            one_base = []
            # Координаты баз переводим из формата [[широта, широта][долгота, долгота]] в формат [[широта, долгота]...[широта, долгота]]
            for base in range(len(bases)):
                one_base.append(bases[base][coord])
            # В зависисмотсти от выбранного пользователем режима выбираем фрагмент ссылки, отвечающий за ограничения в маршруте
            if mode == 1:
                variable = '&routes%5Bavoid%5D=unpaved%2Cpoor_condition'
            elif mode == 2:
                variable = ''
            elif mode == 3:
                variable = '&routes%5Bavoid%5D=unpaved%2Cpoor_condition%2Ctolls'
            elif mode == 4:
                variable = '&routes%5Bavoid%5D=tolls'
            else:
                variable = '&routes%5Bavoid%5D=unpaved%2Cpoor_condition'
            # Формируем остальные части ссылки
            # Часть, отвечающая за ограничения и координаты, ка которых будет открываться для обзора карта
            mode_part = f"?ll={one_base[1]}%2C{one_base[0]}&mode=routes{variable}"
            # Часть, отвечающая за стартовую и конечную точки маршрута
            coordinates_part = f"&rtext={one_base[0]}%2C{one_base[1]}~{coordinates[azs][0]}%2C{coordinates[azs][1]}&rtm=atm&rtt=auto&ruri=~&z=8"
            full_link = const + mode_part + coordinates_part
            print(full_link)
            # Присоединяем готовую ссылку к одной из будущих строк матрицы ссылок
            link_temp.append(full_link)
            step += 1
            print('Поиск маршрутов, шаг номер ', step)
            try:
                # Создаем браузерную сессию, в которую будем отправлять запрос по нашей ссылке
                session = HTMLSession()
                # отправляем запрос, получаем ответ
                response = session.get(full_link, headers=headers)
                # Ждем, пока javascript прогрузится на странице и у нас будет полный набор нужных данных
                response.html.render(sleep=8, timeout=30)
            # Если в подключении возникла ошибка, пробуем снова после небольшого ожидания
            except Exception as error:
                print(error)
                print('--Ошибка в подключении, пробуем снова--')
                time.sleep(15)
                session = HTMLSession()
                response = session.get(full_link, headers=headers)
                response.html.render(sleep=8, timeout=30)
            # Если не получилось снова, то пропускаем эту ссылку
            finally:
                pass
            # Извлекаем из полученной прогруженной html-страницы нужные элементы с искомыми маршрутами
            soup_route = bs4.BeautifulSoup(response.html.html, 'html.parser')
            route = soup_route.findAll("div", {"class": "auto-route-snippet-view__route-subtitle"})
            # Из bs4-html-объекта в str
            route = str(route)
            print(route)
            # Пробуем найти минимальный маршрут, используя функцию поиска всех посчитанных и отрендеренных маршрутов
            try:
                route_str = parse_km_m(route)
                route_str = min(route_str)
                # Присоединяем полученый маршрут к одной из строк будущей матрицы
                all_bases_to_azs.append(route_str)
                print('Success in normal search')
            except:
                # Если не получилось найти маршрут, значит либо яндекс выдал капчу, либо произошла ошибка при
                # подключении, либо произошла ошибка при построении маршрута, то есть м.б. маршрут невозможен
                # В таком случае либо проверяем наличие ошибки построения маршрута, либо после небольшого ожидания
                # отправляем запрос снова, пытаясь таким образом бобйти капчу или ошибку в подключении
                try:
                    # Ищем сообщение об ошибке в построении
                    soup_route = bs4.BeautifulSoup(response.html.html, 'html.parser')
                    soup_error = soup_route.findAll("div", {"class": "route-error-view__text"})
                    # Если сообщение есть - добавляем информацию в матрицу
                    if soup_error != []:
                        all_bases_to_azs.append("route error")
                        print('Error in normal search')
                    # Иначе же пробуем снова и снова отправлять запросы с интервалом в 20 секунд в течение 3х минут
                    else:
                        time_end = time.time() + 60 * 3
                        # Пока ответ не будет получен или таймер не истечет, пробуем
                        while route_str == [] and time.time() < time_end:
                            time.sleep(20)
                            # Создаем новую сессию, в которой буем пытаться добиться нормального ответа
                            session_new = HTMLSession()
                            response = session_new.get(full_link, headers=headers)
                            response.html.render(sleep=8, timeout=30)
                            soup_route = bs4.BeautifulSoup(response.html.html, 'html.parser')
                            route = soup_route.findAll("div", {"class": "auto-route-snippet-view__route-subtitle"})
                            route = str(route)
                            route_str = parse_km_m(route)
                            if route_str == []:
                                soup_route = bs4.BeautifulSoup(response.html.html, 'html.parser')
                                soup_error = soup_route.findAll("div", {"class": "route-error-view__text"})
                                if soup_error != []:
                                    route_str.append('route_error')
                            # Закрываем новую сессию, чтобы не перегружать оперативную память и процессор
                            session_new.close()
                        # Есть результат при поиске, и он не ошибка - добавляем в матрицу
                        if route_str != [] and route_str != ['route_error']:
                            route_str = min(route_str)
                            all_bases_to_azs.append(route_str)
                            print('Success in additional search')
                        # Если в новой сессии выдало сообщение об ошибке в маршруте, добавляем его в матрицу маршрутов
                        elif route_str == ['route_error']:
                            all_bases_to_azs.append(route_str)
                            print('Error in additional search')
                        # Обработка иных исходов, какая-то ещё ошибка
                        else:
                            all_bases_to_azs.append("unknown_error")
                            print('Unknown error occurs')
                # Обработка неизвестных ошибок, которые могут возникнуть
                except Exception as exception:
                    all_bases_to_azs.append("unknown_error_in algorithm: " + str(exception))
                    print('Unknown exception occurs')
                    print(exception)
            # Закрытие сессии, чтобы не оставлять открытым процесс, занимающий оперативку и мощности процессора
            session.close()
        # К матрицам присоединяем готовые сформироавнные строки
        all_routes.append(all_bases_to_azs)
        link_matrix.append(link_temp)
    print('Finish')
    return all_routes, link_matrix


def parse_for_names(coordinates):
    # Поиск названий точек, на которых предположительно располагаются АЗС или оптовики
    # Заголовки, с которыми точно работает код. Нужны при отправке запросов яндексу
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
        'Connection': 'keep-alive',
        'Host': 'market.yandex.ru',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.61 Safari/537.36',
    }
    # Постоянная часть ссылки
    const = "https://yandex.ru/maps/"
    # Список ссылок на точки
    urls = []
    # Список названий найденных точек
    yandex_names = []
    # Счетчик шагов
    step = 0
    for azs in range(len(coordinates)):
        # Часть ссылки отвечающая за режим поиска
        mode = f"?ll={coordinates[azs][1]}%2C{coordinates[azs][0]}&mode=search&"
        # Часть ссылки отвечающая за координаты точки поиска
        coordinates_part = f"sll={coordinates[azs][1]}%2C{coordinates[azs][0]}&text={coordinates[azs][0]}%2C{coordinates[azs][1]}&z=15"
        full_link = const + mode + coordinates_part
        print(full_link)
        urls.append(full_link)
        step += 1
        print('Поиск локаций, шаг номер ', step)
        # Пробуем создать сессию, отправить запрос и получить название точки на координатах
        try:
            session = HTMLSession()
            response = session.get(full_link, headers=headers)
            response.html.render(sleep=8, timeout=30)
            soup_location = bs4.BeautifulSoup(response.html.html, 'html.parser')
            # Точки на странице могут отображаться двумя разными классами элементов, поэтому
            # сначала ищем названия в одном классе элементов, если пусто - то в другом
            location = str(soup_location.find("h1", {"class": "card-title-view__title"}))
            description = str(soup_location.find("div", {"class": "toponym-card-title-view__description"}))
            if location == 'None' or description == 'None':
                location = str(soup_location.find("div", {"class": "search-snippet-view__title"}))
                description = str(soup_location.find("div", {"class": "search-snippet-view__description"}))
                location_str = re.findall(r'<div class="search-snippet-view__title">(.*)</div>', location)
                description_str = re.findall(r'<div class="search-snippet-view__description">(.*)</div>', description)
            else:
                location_str = re.findall(r'<h1 class="card-title-view__title" itemprop="name">(.*)</h1>', location)
                description_str = re.findall(r'<div class="toponym-card-title-view__description">(.*)</div>', description)
            print(location_str, description_str)
            # Присоединяем полученное имя к списку имён
            if location_str != [] and description_str != []:
                yandex_names.append(location_str[0] + ', ' + description_str[0])
            # Обработка ошибок
            else:
                yandex_names.append('error')
            time.sleep(5)
            # Закрываем сессию, чтобы не перегружать устройство
            session.close()
        # Слабая, но обработка исключений
        except:
            yandex_names.append('Unknown exception here')
    return urls, yandex_names


def saving_data(coordinates, urls, names, yandex_names, all_routes, link_matrix, mode):
    # Сохраняем все полученные результаты
    # В зависимости от выбранного режима поиска формируем сообщеине
    print(os.path.abspath(os.path.curdir))
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Маршруты'
        wb.save(filename=f'output_routes_{str(datetime.date.today())}_mode_{mode}.xlsx')
    except:
        pass
    if mode == 1:
        text = '1: Поиск не избегая платных дорог, грунтовые избегаются'
    elif mode == 2:
        text = '2: Поиск не избегая платных дорог, грунтовые не будут избегаться'
    elif mode == 3:
        text = '3: Поиск избегая платных дорог, грунтовые избегаются'
    elif mode == 4:
        text = '4: Поиск избегая платных дорог, грунтовые не будут избегаться'
    else:
        text = '1: Поиск не избегая платных дорог, грунтовые избегаются'
    # Список имён приходит как список списков. Переводим его в обычный список с элементами
    names = list(flatten(names))
    # Список координат добавляем в итоговый файл, но перед этим транспонируем
    coordinates = np.array(coordinates).transpose().tolist()
    # Списки, нужные для формирования словаря в дальнейшем
    data = [coordinates[0], coordinates[1], urls, names, yandex_names]
    columns = ['Широта', 'Долгота', 'Ссылка', 'Исходные имена', 'Полученное имя']
    # Транспонируем матрицы маршрутов и ссылок, чтобы в одном вложенном списке хранились маршруты
    # от одной нефтебазы ко всем АЗС и оптовикам
    all_routes = np.array(all_routes).transpose().tolist()
    link_matrix = np.array(link_matrix).transpose().tolist()
    # Присоединяем к спискам с данными и именами маршруты от нефтебазы до всех АЗС
    for i in range(len(all_routes)):
        data.append(all_routes[i])
        columns.append('НБ' + str(i+1))
    # Две пустых колонки, которые отделять матрицы маршрутов и ссылок
    empty = np.empty((2, len(all_routes[0])))
    empty[:] = np.nan
    columns.append('Режим')
    columns.append('Пусто')
    empty = empty.tolist()
    # К пустым колонкам присоединяем информацию о режиме поиска
    empty[0][0] = text
    data.append(empty[0])
    data.append(empty[1])
    # Присоединяем к спискам с данными матрицу ссылок
    for i in range(len(link_matrix)):
        data.append(link_matrix[i])
        columns.append('НБ' + str(i+1) + '_url')
    # Формируем словарь, чтобы его перекастовать в DataFrame-объект и потом его сохранить в файл
    dictionary = {}
    for i in range(len(data)):
        dictionary.update({f'{columns[i]}': data[i]})
    # Перекастовываем
    data = pd.DataFrame(data=dictionary)
    # Пробуем сохранить
    try:
        with pd.ExcelWriter(f'output_routes_{str(datetime.date.today())}_mode_{mode}.xlsx', engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
            data.to_excel(writer, sheet_name="Маршруты")
    # Если не получилось сохранить полностью DataFrame-объект с данными, то сохраняем все данные на отдельные листы,
    # чтобы не потерять данные и понять, где возникла ошибка
    except Exception as error:
        print('error:', error)
        for i in range(len(columns)):
            try:
                with pd.ExcelWriter(f'output_routes_{str(datetime.date)}_mode_{mode}.xlsx', engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
                    pd.DataFrame(data[i]).to_excel(writer, sheet_name=f"{columns[i]}")
            except:
                pass
    finally:
        print('Программа закончила свою работу')


def partial_saving(data):
    # Функция для сохранения какого-то отдельного типа данных
    data = pd.DataFrame(data)
    with pd.ExcelWriter(f'output_routes_{str(datetime.date)}_mode_{mode}.xlsx', engine='openpyxl', mode='a', if_sheet_exists="replace") as writer:
        data.to_excel(writer, sheet_name=f"Data_partial_{random.randint(1, 1000)}")

# Начало работы с программой, пока пользователь не выберет правильный режим - ничего не делаем. По умолчанию режим 1
while True:
    try:
        mode = int(input('Введите, 1, 2, 3 или 4 в зависимости от того, какой режим нужен для поиска маршрутов (по умолчанию 1):\n'
                 '1: Поиск не избегая платных дорог, грунтовые избегаются\n'
                 '2: Поиск не избегая платных дорог, грунтовые не будут избегаться\n'
                 '3: Поиск избегая платных дорог, грунтовые избегаются\n'
                 '4: Поиск избегая платных дорог, грунтовые не будут избегаться\n')) or 1
        if mode == 1 or 2 or 3 or 4:
            break
        else:
            print('Введено некорректное значение')
    except:
        mode = 1
        break


# Здесь задается путь до файла, который считываем
coordinates, bases, names = read_dataset("Координаты пример.xlsx")
all_routes, link_matrix = parse_for_routes(coordinates, bases, mode)
urls, yandex_names = parse_for_names(coordinates)
saving_data(coordinates, urls, names, yandex_names, all_routes, link_matrix, mode)
# Если что-то отработало с ошибкой, то можно выполнить только одну из функций выше и нужную часть данных сохранить при помощи функции ниже

# partial_saving(yandex_names)
