﻿using Novacode;
using Oracle.DataAccess.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace GBConverter {
    class Converter {
        const int RESULT_COLUMN = 13;

        Dictionary<string, string> Subjects = new Dictionary<string, string>();
        Dictionary<string, string> Confirmations = new Dictionary<string, string>();

        static void _t(String str) {
            System.Diagnostics.Trace.WriteLine(str);
        }

        public void Convert(string fileName, ProgressBar progressBar = null) {
            if (fileName == "") {
                throw new ArgumentException("Не указано имя файла для конвертации");
            }
            // Загружаем справочник субъектов РФ из БД
            if (!this.LoadDictionaries()) {
                return;
            }

            ArrayList RowErrors = new ArrayList();
            ArrayList SimpleAppeals = new ArrayList();
            DocX document;
            try {
                document = DocX.Load(fileName);
            } catch (System.IO.IOException e) {
                MessageBox.Show(String.Format("Не удалось открыть файл {0}.\nВозможно он используется другой программой.\n\nРабота программы прекращена.", fileName),
                    "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            Table appealsTable = document.Tables[0];

            // Добавляем колонку с информацией о выявленных несоответствиях по обращению (результат проверки).
            appealsTable.InsertColumn();
            appealsTable.Rows[0].Cells[RESULT_COLUMN].VerticalAlignment = VerticalAlignment.Top;
            Paragraph p = appealsTable.Rows[0].Cells[RESULT_COLUMN].Paragraphs[0].Append("Выявленные несоответствия данных (для конвертации)");
            p.Alignment = Alignment.center;
            p.Bold();
            bool Step2 = true;
            // Анализируем все строки таблицы.
            for (int rowIndex = 1; rowIndex < appealsTable.Rows.Count; rowIndex++) {
                // checkResult = CheckAppeal(appealsTable.Rows[rowIndex]);
                if (CheckAppeal(appealsTable.Rows[rowIndex], SimpleAppeals, out RowErrors)) {
                    // Проверка обращения прошла успешно.
                } else {
                    // Проверка обращения завершилась с ошибками.
                    Step2 = false;
                    // Записываем результат проверки в правую колонку.
                    foreach (string str in RowErrors) {
                        appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs[0].Append(str);
                    }
                }
                double percent = (double) rowIndex / (appealsTable.Rows.Count - 1) * 100;
                progressBar.Value = System.Convert.ToInt32(percent);
            }

            if (Step2) {
                // Переходим ко второму этапу - запись в БД
                MessageBox.Show("Несоответствий не выявлено.", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } else {
                string filePath = Path.GetDirectoryName(fileName) + @"\Ошибки конвертации " + (DateTime.Now).ToString("yyyy-MM-dd-HH-mm-ss") + ".docx";
                document.SaveAs(filePath);
                System.Diagnostics.Process.Start(filePath);
            }
            document.Dispose();
            // System.Diagnostics.Trace.WriteLine((DateTime.Now).ToString( "yyyy-MM-dd-H-mm-ss" ));
        }
        /// <summary>
        /// <para>Проверяет данные заявки</para>
        /// <para>Возвращает true если данные разобраны без ошибок и false в противном случае.</para>
        /// </summary>
        /// <param name="row">Строка таблицы.</param>
        /// <param name="parsedData">out ArayList с данными либо с ошибками, если разбор завершился неудаче.</param>
        bool CheckAppeal(Row row, ArrayList appeals, out ArrayList errors) {
            bool result = true;
            string cellText = "";
            string tmp;
            errors = new ArrayList();
            ArrayList cellParsedText;
            ArrayList Subjects = new ArrayList();
            ArrayList NumbersAndDates = new ArrayList();
            ArrayList Declarants = new ArrayList();
            //appeals = new ArrayList();
            // Объект для хранения общих данных обращения
            Appeal NewAppeal = new Appeal();
            // Пропускаем первую и последнюю колонку.
            for (int colIndex = 1; colIndex < row.Cells.Count - 1; colIndex++) {
                Cell c = row.Cells[colIndex];
                cellText = "";
                tmp = "";
                // Собираем текст ячейки из всех параграфов в одну переменную.
                for (int i = 0; i < c.Paragraphs.Count; i++) {
                    if (c.Paragraphs[i].Text.Trim() != "") {
                        cellText += c.Paragraphs[i].Text + " ";
                    }
                }
                // Запускаем разбор текста из ячейки.
                switch (colIndex) {
                    // Субъект РФ+
                    case 1:
                        if (ParseSubject(cellText, out cellParsedText)) {
                            // Проверка завершилась успешно - заполняем массив субъектов
                            foreach (string str in cellParsedText) {
                                Subjects.Add(str);
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add("Субъект Российской Федерации: ");
                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedText) {
                                errors[errors.Count - 1] += str + "; ";
                            }
                        }
                        break;
                    // Содержание+
                    case 2:
                        if (ParseContent(cellText, out tmp)) {
                            NewAppeal.content = tmp;
                        } else {
                            result = false;
                            errors.Add("Содержание: " + tmp);
                        }
                        break;
                    // Заявитель+
                    case 3:
                        if (ParseDeclarant(cellText, out cellParsedText)) {
                            // Проверка завершилась успешно - заполняем массив субъектов
                            foreach (string str in cellParsedText) {
                                Declarants.Add(str);
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add("Кем заявлено: ");
                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedText) {
                                errors[errors.Count - 1] += str + "; ";
                            }
                        }
                        break;
                    // Сведения о подтверждении
                    case 4:
                        if (ParseConfirmation(cellText, out tmp)) {
                            NewAppeal.confirmation = tmp;
                        } else {
                            result = false;
                            errors.Add("Сведения о подтверждении: " + tmp);
                        }
                        break;
                    // Приянтые меры
                    case 5:
                        if (ParseMeasures(cellText, out tmp)) {
                            NewAppeal.measures = tmp;
                        } else {
                            result = false;
                            errors.Add("Приянтые меры: " + tmp);
                        }
                        break;
                    // Номер и дата
                    case 6:
                        if (ParseNumberAndDate(cellText, out cellParsedText)) {
                            // Проверка завершилась успешно - заполняем массив субъектов
                            foreach (Tuple<string, string> NumDate in cellParsedText) {
                                NumbersAndDates.Add(NumDate);
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add("Рег. номер и дата: ");
                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedText) {
                                errors[errors.Count - 1] += str + "; ";
                            }
                        }
                        break;
                    // 7 - Уровень выборов
                    // Партия
                    case 8:
                        if (ParseParty(cellText, out tmp)) {
                            NewAppeal.party = tmp;
                        } else {
                            result = false;
                            errors.Add("Партия: " + tmp);
                        }

                        break;
                    // Тип заявителя
                    case 9:
                        if (ParseDeclarantType(cellText, out tmp)) {
                            NewAppeal.declarant_type = tmp;
                        } else {
                            result = false;
                            errors.Add("Тип заявителя: " + tmp);
                        }

                        break;
                    // Тематика
                    case 10:
                        if (ParseTheme(cellText, out tmp)) {
                            NewAppeal.theme = tmp;
                        } else {
                            result = false;
                            errors.Add("Тематика: " + tmp);
                        }
                        break;
                    // 11 - +
                    // Исполнитель
                    case 12:
                        if (ParseExecutor(cellText, out tmp)) {
                            NewAppeal.executor_id = tmp;
                        } else {
                            result = false;
                            errors.Add("Исполнитель: " + tmp);
                        }
                        break;
                }
            }

            // Создаём "простые" обращения
            foreach (string SubjCode in Subjects) {
                foreach (Tuple<string, string> NumDate in NumbersAndDates) {
                    NewAppeal.subjcode = SubjCode;
                    NewAppeal.numb = NumDate.Item1;
                    NewAppeal.f_date = NumDate.Item2;
                    appeals.Add(NewAppeal);
                }
            }
            return result;
        }
        bool ParseSubject(string inputData, out ArrayList resultData) {
            bool result = false;
            resultData = new ArrayList();
            // Удаляем лишние символы.
            inputData = PrepareRawData(inputData);
            // Разделяем текст на части. В Words будут записаны названия субъектов.
            char[] Separatos = { ';' };
            string[] Words = inputData.Split(Separatos, StringSplitOptions.RemoveEmptyEntries);

            // Ищем код каждого субъекта по его названию.
            try {
                foreach (string SubjectName in Words) {
                    KeyValuePair<string, string> Subj = Subjects.First(s => s.Key == SubjectName.Trim());
                    resultData.Add(Subj.Value);
                }
                result = true;
            } catch (System.InvalidOperationException) {
                resultData.Clear();
                resultData.Add("Наименование субъекта РФ не найдено в справочнике \"Субъекты РФ\"");
            }

            return result;
        }
        bool ParseContent(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                parsed = PrepareRawData(data);
                result = true;
            }
            return result;
        }
        bool ParseDeclarant(string data, out ArrayList resultData) {
            bool result = false;
            resultData = new ArrayList();
            // Заглушка
            resultData.Add("Какой-то заявитель");
            result = true;
            return result;
        }
        bool ParseConfirmation(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                // Заглушка
                parsed = "Какое-то подтверждение";
                result = true;
            }
            return result;
        }
        bool ParseMeasures(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                parsed = PrepareRawData(data);
                result = true;
            }
            return result;
        }
        bool ParseNumberAndDate(string appealNumber, out ArrayList appealDate) {
            bool result = true;
            appealDate = new ArrayList();
            //string sPattern = @"^\d{3}-\d{3}-\d{4}$";
            Tuple<string, string> NumDate;
            NumDate = Tuple.Create("1111", "12.12.2014");
            appealDate.Add(NumDate);
            NumDate = Tuple.Create("2222", "05.09.2015");
            appealDate.Add(NumDate);
            NumDate = Tuple.Create("3333", "11.10.2015");
            appealDate.Add(NumDate);
            //appealDate.Add

            return result;
        }
        bool ParseParty(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                // Заглушка
                parsed = "Какая-то партия";
                result = true;
            }
            return result;
        }
        bool ParseDeclarantType(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                // Заглушка
                parsed = "Какой-то заявитель";
                result = true;
            }
            return result;
        }
        bool ParseExecutor(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                // Заглушка
                parsed = "Какой-то исполнитель";
                result = true;
            }
            return result;
        }
        bool ParseTheme(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле";
            } else {
                // Заглушка
                parsed = "Так себе тема";
                result = true;
            }
            return result;
        }
        string PrepareRawData(String data) {
            // Удаляет лишние символы
            string result = Regex.Replace(data, @"\s+", " ");
            return result;
        }
        bool LoadDictionaries() {

            // UNCOMMENT!
            string oradb = "Data Source=RA00C000;User Id=voshod;Password=voshod;";
            OracleConnection conn = new OracleConnection(oradb);
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            OracleDataReader dr = null;
            cmd.CommandType = CommandType.Text;
            // Субъекты.
            cmd.CommandText = "select namate, subjcod from ate_history where prsubj='1' and datedel is null";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                conn.Dispose();
                return false;
            }

            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    Subjects.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                }
                //subjects.FirstOrDefault()
            }


            dr.Dispose();
            cmd.Dispose();
            conn.Dispose();


            // DB emulation :)
            /*
            Confirmations.Add("Нарушение не подтвердилось", "1");
            Confirmations.Add("Нарушение подтвердилось", "2");
            Confirmations.Add("Нарушение подтвердилось частично", "3");

            Subjects.Add("Республика Адыгея", "01");
            Subjects.Add("Республика Алтай", "04");
            Subjects.Add("Республика Башкортостан", "02");
            Subjects.Add("Республика Бурятия", "03");
            Subjects.Add("Республика Дагестан", "05");
            Subjects.Add("Республика Ингушетия", "06");
            Subjects.Add("Кабардино-Балкарская республика", "07");
            Subjects.Add("Республика Калмыкия", "08");
            Subjects.Add("Карачаево-Черкесская республика", "09");
            Subjects.Add("Республика Карелия", "10");
            Subjects.Add("Республика Коми", "11");
            Subjects.Add("Республика Крым", "91");
            Subjects.Add("Республика Марий Эл", "12");
            Subjects.Add("Республика Мордовия", "13");
            Subjects.Add("Республика Саха (Якутия)", "14");
            Subjects.Add("Республика Северная Осетия — Алания", "15");
            Subjects.Add("Республика Татарстан", "16");
            Subjects.Add("Республика Тыва", "17");
            Subjects.Add("Удмуртская республика", "18");
            Subjects.Add("Республика Хакасия", "19");
            Subjects.Add("Чеченская республика", "20");
            Subjects.Add("Чувашская республика", "21");
            Subjects.Add("Алтайский край", "22");
            Subjects.Add("Краснодарский край", "23");
            Subjects.Add("Красноярский край", "24");
            Subjects.Add("Приморский край", "25");
            Subjects.Add("Ставропольский край", "26");
            Subjects.Add("Хабаровский край", "27");
            Subjects.Add("Амурская область", "28");
            Subjects.Add("Архангельская область", "29");
            Subjects.Add("Астраханская область", "30");
            Subjects.Add("Белгородская область", "31");
            Subjects.Add("Брянская область", "32");
            Subjects.Add("Владимирская область", "33");
            Subjects.Add("Волгоградская область", "34");
            Subjects.Add("Вологодская область", "35");
            Subjects.Add("Воронежская область", "36");
            Subjects.Add("Ивановская область", "37");
            Subjects.Add("Иркутская область", "38");
            Subjects.Add("Калининградская область", "39");
            Subjects.Add("Калужская область", "40");
            Subjects.Add("Кемеровская область", "42");
            Subjects.Add("Кировская область", "43");
            Subjects.Add("Костромская область", "44");
            Subjects.Add("Курганская область", "45");
            Subjects.Add("Курская область", "46");
            Subjects.Add("Ленинградская область", "47");
            Subjects.Add("Липецкая область", "48");
            Subjects.Add("Магаданская область", "49");
            Subjects.Add("Московская область", "50");
            Subjects.Add("Мурманская область", "51");
            Subjects.Add("Нижегородская область", "52");
            Subjects.Add("Новгородская область", "53");
            Subjects.Add("Новосибирская область", "54");
            Subjects.Add("Омская область", "55");
            Subjects.Add("Оренбургская область", "56");
            Subjects.Add("Орловская область", "57");
            Subjects.Add("Пензенская область", "58");
            Subjects.Add("Псковская область", "60");
            Subjects.Add("Ростовская область", "61");
            Subjects.Add("Рязанская область", "62");
            Subjects.Add("Самарская область", "63");
            Subjects.Add("Саратовская область", "64");
            Subjects.Add("Сахалинская область", "65");
            Subjects.Add("Свердловская область", "66");
            Subjects.Add("Смоленская область", "67");
            Subjects.Add("Тамбовская область", "68");
            Subjects.Add("Тверская область", "69");
            Subjects.Add("Томская область", "70");
            Subjects.Add("Тульская область", "71");
            Subjects.Add("Тюменская область", "72");
            Subjects.Add("Ульяновская область", "73");
            Subjects.Add("Челябинская область", "74");
            Subjects.Add("Ярославская область", "76");
            Subjects.Add("Москва", "77");
            Subjects.Add("Санкт-Петербург", "78");
            Subjects.Add("Севастополь", "92");
            Subjects.Add("Еврейская автономная область", "79");
            Subjects.Add("Ненецкий автономный округ", "83");
            Subjects.Add("Ханты-Мансийский автономный округ - Югра", "86");
            Subjects.Add("Чукотский автономный округ", "87");
            Subjects.Add("Ямало-Ненецкий автономный округ", "89");
            */
            return true;
        }
    }
}