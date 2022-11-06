# Парсинг данных о длине маршрутов из Яндекс Карт
## Описание проекта
Программа предназначена для поиска маршрутов между точками, заданными в виде координат, а также для определения того, что за объект находится на выбранных координатах. Программа составляет ссылку на маршрут между точками, отправляет запрос по этой ссылке, и, получив ответ, ищет на странице число километров между точками. После этого программа составляет ссылку на одну из точек, отправляет запрос и ищет в полученном ответе адрес объекта. Итоговый результат формируется в виде матрицы расстояний между точками и матрицы ссылок на маршруты, матрицы сохраняются в excel-файл.
Запросы отправляются при помощи requests_html. Чтобы избежать возникновение капчи, запросы отправляются с интервалами в 5-30 секунд. Если не удалось избежать капчи, то запросы будут в течение некторого времени повторно отправляться до получения ответа, при неудаче вместо ответа будет получено сообщение о неудаче и процесс продолжится. 

## Требования для установки
- Python версии 3.x
- Установка пакетов, перечисленных в requirements.txt

## Как использовать программу
Запуск производится из файла Yan_EMIS.py (в терминале: '''py Yan_EMIS.py''').
В проекте присутствует тестовый файл, на котором можно опробовать программу.
При первом запуске программы она может вместо отправки запросов начать устанавливать chromium, так как он нужен для создания сессий. 
Путь до файла, с которого будут считывться данные нужно указывать в функции read_dataset() (395 стр. кода).
Внутри этой функции можно вручную корректировать то, какие столбцы и строки файла считывать.
Для выбора столбцов для считывания нужно менять параметр usecols=, для пропуска строчек, например пустых, указывается параметр skiprows=, число считываемых строчек можно указать параметром nrows=.
Во время работы программа может писать предупреждения вида "Удалённый хост принудительно разорвал существующее подключение",
что не является проблемой, происходит закрытие предыдущих сессий.
