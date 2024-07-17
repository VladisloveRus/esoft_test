import csv
import datetime
import re

import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt

# Имя анализируемого файла:
INPUT_FILE_NAME = "Экспозиция ТДСК с 01.07.2023 по 31.12.2023.csv"
# Ограничение выборки по ранней дате:
DATE_EARLIEST_LIMIT = "2023-07-01"  # str, формат ГГГГ.ММ.ДД
# Ограничение выборки по поздней дате:
DATE_LATEST_LIMIT = "2023-12-31"  # str, формат ГГГГ.ММ.ДД
# Текстовое обозначение для отсутвия объектов
IDENTIFICATOR_NO_DATA = "No data"
# Режим дебагинга:
DEBUG = False


now_date = str(datetime.date.today())


def get_date(string: str):
    """Перевод даты из str в datetime"""
    search = re.search(r"(?P<year>.*)\-(?P<month>.*)\-(?P<day>..)", string)
    if not search:
        if DEBUG:
            print(
                f"Проблема на этапе определения даты. Текст ячейки: {string}"
            )
        raise ValueError(
            "Следующие данные не соответствуют требованиям регулярного "
            f"выражения: {string}"
        )
    result = datetime.date(
        int(search.group("year")),
        int(search.group("month")),
        int(search.group("day")),
    )
    return result


# Имя результирующего файла:
OUTPUT_FILE_NAME = (
    f"./Выборка от {now_date}, период {DATE_EARLIEST_LIMIT} - "
    f"{DATE_LATEST_LIMIT}.xlsx"
)

begin_date_limit = get_date(DATE_EARLIEST_LIMIT)
end_date_limit = get_date(DATE_LATEST_LIMIT)


def create_base_table(begin_limit, end_limit):
    """Заполнение таблицы необходимыми датами"""
    date_delta = (end_limit - begin_limit).days
    current_date = begin_limit
    table = dict()
    while date_delta >= 0:
        table[current_date] = IDENTIFICATOR_NO_DATA
        current_date += datetime.timedelta(days=1)
        date_delta -= 1
    return table


def create_month_table(begin_limit, end_limit):
    """Заполнение таблицы необходимыми месяцами"""
    date_delta = (
        end_limit.month + (end_limit.year - begin_limit.year) * 12
    ) - begin_limit.month
    current_month_year = [begin_limit.month, begin_limit.year]
    table = dict()
    while date_delta >= 0:
        month_year_to_str = list()
        for item in current_month_year:
            month_year_to_str.append(str(item))
        month_year_record = ", ".join(month_year_to_str)
        table[month_year_record] = IDENTIFICATOR_NO_DATA
        if (current_month_year[0] + 1) > 12:
            current_month_year[0] = 0
            current_month_year[1] += 1
        current_month_year[0] += 1
        date_delta -= 1
    return table


def create_excel_file(date_list: list, address_list: list, amount_list: list):
    """Создание xlsx-файла с данными"""
    df = pd.DataFrame(
        {
            "Дата": date_list,
            "Корпус": address_list,
            "Кол-во активных квартир": amount_list,
        }
    )
    df.to_excel(
        OUTPUT_FILE_NAME,
        sheet_name=f"{DATE_EARLIEST_LIMIT} - {DATE_LATEST_LIMIT}",
        index=False,
        engine="openpyxl",
    )
    return correcting_width()


def correcting_width():
    """Форматирование ширины и высоты столбцов таблицы"""
    wb = load_workbook(OUTPUT_FILE_NAME)
    ws = wb.active
    ws.title = "RowColDimension"
    ws.column_dimensions["A"].width = 15
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 25
    return wb.save(OUTPUT_FILE_NAME)


def collect_active_objects():
    """Создание сводной таблицы с общим количеством активных объектов за каждый день рассматриваемого периода по каждому из корпусов"""
    table = create_base_table(begin_date_limit, end_date_limit)
    with open(
        INPUT_FILE_NAME,
        encoding="utf-8",
        newline="",
    ) as csvfile:
        reader = csv.reader(csvfile, delimiter="\t")
        for row in reader:
            search_address = re.search(r"(?P<address>.*)\, подъезд", row[4])

            if not search_address:
                if DEBUG:
                    print(f"Проблема c получением адреса в строке {row}")
                continue
            address = search_address.group("address")
            result_address = address.replace(
                "  ", " "
            )  # Перфекционизм наше всё

            begin_date = get_date(row[-2])
            if (begin_date - begin_date_limit).days < 0:
                begin_date = begin_date_limit
            end_date = get_date(row[-1])
            if (end_date_limit - end_date).days < 0:
                end_date = end_date_limit
            date_delta = (end_date - begin_date).days
            current_date = begin_date

            while date_delta >= 0:
                date_delta -= 1
                if table[current_date] == IDENTIFICATOR_NO_DATA:
                    table[current_date] = {}
                try:
                    table[current_date][result_address] = (
                        table[current_date][result_address] + 1
                    )
                    if DEBUG:
                        print(
                            "Увеличено количество объектов на дату "
                            f"{current_date} по адресу {result_address} до "
                            f"{table[current_date][result_address]}"
                        )
                except KeyError:
                    table[current_date][result_address] = 1
                    if DEBUG:
                        print(
                            f"Добавлен новый адрес {result_address} на дату "
                            f"{current_date}"
                        )
                current_date = current_date + datetime.timedelta(days=1)

    keys_list = table.keys()

    date_list = []
    address_list = []
    amount_list = []
    for date in keys_list:
        try:
            address_amount_list = table[date].keys()
        except AttributeError:
            table[date] = {IDENTIFICATOR_NO_DATA: IDENTIFICATOR_NO_DATA}
            address_amount_list = table[date].keys()
        for address in address_amount_list:
            date_list.append(date)
            address_list.append(address)
            amount_list.append(table[date][address])
            if DEBUG:
                print(
                    f"На {date} по адресу {address} количество объектов: "
                    f"{table[date][address]}"
                )
    return create_excel_file(date_list, address_list, amount_list)


def graph_output():
    """Вывод графика по месячному количеству активных объектов в разрезе комнатности"""
    table = create_month_table(begin_date_limit, end_date_limit)
    with open(
        INPUT_FILE_NAME,
        encoding="utf-8",
        newline="",
    ) as csvfile:
        reader = csv.reader(csvfile, delimiter="\t")
        for row in reader:
            try:
                number_of_room = int(row[10])
            except ValueError:
                continue
            begin_date = get_date(row[-2])
            if (begin_date - begin_date_limit).days < 0:
                begin_date = begin_date_limit
            end_date = get_date(row[-1])
            if (end_date_limit - end_date).days < 0:
                end_date = end_date_limit
            begin_month = begin_date.month
            begin_year = begin_date.year
            current_month_year = [begin_month, begin_year]
            end_month = end_date.month
            end_year = end_date.year
            date_delta = (
                end_month + (end_year - begin_year) * 12
            ) - begin_month

            while date_delta >= 0:
                month_year_to_str = list()
                for item in current_month_year:
                    month_year_to_str.append(str(item))
                month_year_record = ", ".join(month_year_to_str)
                date_delta -= 1
                if table[month_year_record] == IDENTIFICATOR_NO_DATA:
                    table[month_year_record] = {}
                try:
                    table[month_year_record][number_of_room] = (
                        table[month_year_record][number_of_room] + 1
                    )
                    if DEBUG:
                        print(
                            "Увеличено количество объектов на месяц "
                            f"{month_year_record} по количеству комнат "
                            f"{number_of_room} до "
                            f"{table[month_year_record][number_of_room]}"
                        )
                except KeyError:
                    table[month_year_record][number_of_room] = 1
                    if DEBUG:
                        print(
                            f"Добавлен новое количество комнат "
                            f"{number_of_room} на месяц {month_year_record}"
                        )
                if (current_month_year[0] + 1) > 12:
                    current_month_year[0] = 0
                    current_month_year[1] += 1
                current_month_year[0] += 1
    keys_list = table.keys()
    for month_year in keys_list:
        table[month_year] = dict(sorted(table[month_year].items()))

    month_year_list = []
    rooms_list = []
    amount_list = []
    for month_year in keys_list:
        try:
            rooms_amount_list = table[month_year].keys()
        except AttributeError:
            table[month_year] = {IDENTIFICATOR_NO_DATA: IDENTIFICATOR_NO_DATA}
            rooms_amount_list = table[month_year].keys()
        for room in rooms_amount_list:
            month_year_list.append(month_year)
            rooms_list.append(room)
            amount_list.append(table[month_year][room])
            if DEBUG:
                print(
                    f"На месяц {month_year} год по количеству комнат {room} "
                    f"количество объектов: {table[month_year][room]}"
                )
    fig, ax = plt.subplots()
    x = month_year_list[::3]
    y1 = amount_list[::3]
    y2 = amount_list[1::3]
    y3 = amount_list[2::3]
    ax.plot(x, y1, label="Однокомнатные")
    ax.plot(x, y2, label="Двухкомнатные")
    ax.plot(x, y3, label="Трехкомнатные")
    ax.set_title(
        "Месячное количество активных объектов в разрезе комнатности"
    )
    ax.set_xlabel("Месяц, год")
    ax.set_ylabel("Активные объекты, шт.")
    ax.legend(loc="best")
    plt.grid()
    return plt.show()


def main():
    """Основная функция приложения"""
    collect_active_objects()
    graph_output()


if __name__ == "__main__":
    main()
