# esoft_test
## Тестовое задание, выполненное для Esoft

### Инструкция по запуску:

* Создаём и активируем виртуальное окружение
```python -m venv venv```
```. venv/Scripts/activate```

* Устанавливаем зависимости из файла requirements.txt
```pip install -r requirements.txt```

* Для запуска программы используем команду
```python main.py```

### Что происходит после запуска?

Программа считывает файл из корневой папки проекта, название которого указано в константе INPUT_FILE_NAME в файле main.py
В корневой папке проекта появляется файл с названием по форме "Выборка от {Сегодня}, период {Начало периода} - {Конец периода}.xlsx"
Файл содержит информацию о количестве активных объектов по каждому встречающемуся адресу на каждый день в рамках указанного периода
На экран выводится график по месячному количеству активных объектов в разрезе комнатности в рамках указанного периода


Для настройки диапазона получаемой выборки нужно поменять значения DATE_EARLIEST_LIMIT и DATE_LATEST_LIMIT в файле main.py

Результат выполнения программы можно наблюдать в директории example. Там же находится описание полученного графика (описание не генерируется автоматически, анализ проведён мной вручную)

#### TODO:
* Привести README.md в красивый вид, добавить технологии
* Автоматизировать получение максимального числа комнат при формировании графика (всё сломается, если комнат будет больше трёх)
* Вынести константы для настройки работы приложения в файл settings.py
* Провести рефакторинг проекта (убрать нагромождения кода, привести к принципу DRY)

#### Технологии:
*Python
*Pandas
*openpyxl
*matplotlib

#### Автор: [VladisloveRus](https://github.com/VladisloveRus "github.com/VladisloveRus")