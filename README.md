# Конвертер "Зелёной книги" АКРИКО

Для сборки данного проекта Вам понадобится Visual Studio Express 2013 for Desktop и библиотеки [DocX](http://docx.codeplex.com/) и Oracle.DataAccess (включены)

Что нужно сделать:
- Построчный анализ корректности данных
	+ Субъект Российской Федерации
	+ Содержание
	+ Кем заявлено
		заполнить справочник
	+ Сведения о подтверждении	
	+ Принятые меры	
	Рег. номер и дата	
	у	
	п	
	+ з	
	+ т	
	+	
	+ Исполнитель

+ Добавление информации об ошибках в таблицу
- Сохранение корректных "разобранных" данных в массиве
- Добавление обращений в БД
	AddAppeal - добавление одного обращения
	Если в строке таблицы несколько субъектов/номеров - добавляется несколько обращений
	Количество заявителей в строке таблицы не влияет на количество добавляемых обращений, но влияет на количество записей в appeal_multi
	
	
	CheckAppeal(Row row, out ArrayList appeals, out string error) должна возвращать список обращений (ArrayList <Appeal>*)
	* Класс Appeal должен содержать поле appeal_multi ArrayList<string[]>
		в результате разбора каждой строки таблицы должны появиться массивы субъектов и номеров+дат
		во вложенном цикле необходимо сформировать "простые" обращения
		для "родительского" обращения (либо если в строке талицы указан 1 субъект РФ) необходимо внести записи в appeal_multi
		
+ Формирование отчёта о незагруженных обращениях
+ Графический интерфейс
- Журнал загрузки для "отката" загруженных обращений
+ Изменение цвета текста в ячейках с ошибкой RGB(204, 0, 153)
+ Изменить алгоритм поиска субъекта
+ Сделать поле "П" необязательным
+ Не учитывать правую колонку