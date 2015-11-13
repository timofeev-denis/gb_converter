# Конвертер "Зелёной книги" АКРИКО

Для сборки данного проекта Вам понадобится Visual Studio Express 2013 for Desktop и библиотеки [DocX](http://docx.codeplex.com/) и Oracle.DataAccess (включены)

Что нужно сделать:

+ Добавление информации об ошибках в таблицу
+ Сохранение корректных "разобранных" данных в массиве

- Множественное создание заявок
- Запись в log_appeal
- Формирование отчёта о незагруженных обращениях (дублирование номера и даты обращения и пр.)
- Добавление обращений в БД
	+ Если в строке таблицы несколько субъектов/номеров - добавляется несколько обращений
	+ Количество заявителей в строке таблицы не влияет на количество добавляемых обращений, но влияет на количество записей в appeal_multi и cat_declarants
	+ CheckAppeal(Row row, out ArrayList appeals, out string error) должна возвращать список обращений (ArrayList <Appeal>*)
	+ Структура Appeal должен содержать поле multi ArrayList<string[]>
		+ в результате разбора каждой строки таблицы должны появиться массивы субъектов и номеров+дат
		+ во вложенном цикле необходимо сформировать "простые" обращения
		+ внести записи в appeal_multi

+ Журнал загрузки для "отката" загруженных обращений
+ Не добавлять заявителя, если он указан в обращении повторно
+ Графический интерфейс
+ Изменение цвета текста в ячейках с ошибкой RGB(204, 0, 153)
+ Изменить алгоритм поиска субъекта
+ Сделать поле "П" необязательным
+ Не учитывать правую колонку