using Novacode;
using Oracle.DataAccess.Client;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace GBConverter {
    enum MessageType { Header, Text };

    class Converter {
        const int RESULT_COLUMN = 13;

        Dictionary<string, string> Subjects = new Dictionary<string, string>();
        Dictionary<string, string> Confirmations = new Dictionary<string, string>();
        Dictionary<string, string> Themes = new Dictionary<string, string>();
        Dictionary<string, string> Parties = new Dictionary<string, string>();
        Dictionary<string, string> DecTypes = new Dictionary<string, string>();
        Dictionary<string, string> Executors = new Dictionary<string, string>();
        List<string[]> Declarants = new System.Collections.Generic.List<string[]>();
        long DeclarantFakeID = 1;

        static void _t(String str) {
            System.Diagnostics.Trace.WriteLine(str);
        }

        public void Convert(string fileName, ProgressBar progressBar = null) {
            if (fileName == "") {
                throw new ArgumentException("Не указано имя файла для конвертации");
            }
            // Загружаем справочники из БД
            if (!this.LoadDictionaries()) {
                return;
            }

            ArrayList RowErrors = new ArrayList();
            ArrayList SimpleAppeals = new ArrayList();
            DocX document;
            try {
                // Открываем файл с "ЗЕлёной книгой".
                document = DocX.Load(fileName);
            } catch (System.IO.IOException e) {
                MessageBox.Show(String.Format("Не удалось открыть файл {0}.\nВозможно он используется другой программой.\n\nРабота программы прекращена.", fileName),
                    "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            Table appealsTable = document.Tables[0];

            // Добавляем колонку с информацией о выявленных несоответствиях по обращению (результат проверки).
            if (appealsTable.ColumnCount > RESULT_COLUMN) {
                appealsTable.RemoveColumn(RESULT_COLUMN);
            }
            appealsTable.InsertColumn();
            appealsTable.Rows[0].Cells[RESULT_COLUMN].VerticalAlignment = VerticalAlignment.Top;
            Paragraph p = appealsTable.Rows[0].Cells[RESULT_COLUMN].Paragraphs[0].Append("Выявленные несоответствия данных (для конвертации)");
            p.Alignment = Alignment.center;
            p.Bold();
            p.FontSize(10);
            bool Step2 = true;

            // Анализируем все строки таблицы.
            for (int rowIndex = 1; rowIndex < appealsTable.Rows.Count; rowIndex++) {
                if (CheckAppeal(appealsTable.Rows[rowIndex], SimpleAppeals, out RowErrors)) {
                    // Проверка обращения прошла успешно.
                } else {
                    // Проверка обращения завершилась с ошибками.
                    Step2 = false;
                    bool FirstError = true;
                    // Записываем результат проверки в правую колонку.
                    foreach (ErrorMessage msg in RowErrors) {
                        if (msg.Type == MessageType.Header) {
                            if (FirstError) {
                                p = appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs[0];
                            } else {
                                // Добавляем пустую строку
                                appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].InsertParagraph("");
                                p = appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].InsertParagraph();
                            }
                            if (msg.Column > 0) {
                                Color violetColor = Color.FromArgb(204, 0, 153);
                                //p.Color(violetColor);
                                foreach (Paragraph ColorParagraph in appealsTable.Rows[rowIndex].Cells[msg.Column].Paragraphs) {
                                    ColorParagraph.Color(violetColor);
                                }
                            }
                            p.Append(msg.Message);
                            p.Bold();
                            p.FontSize(10);
                            FirstError = false;
                        } else {
                            int cnt = appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs.Count;
                            appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs[cnt - 1].Append(msg.Message);
                            appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs[cnt - 1].FontSize(10);
                        }
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
            ArrayList cellParsedValues;
            ArrayList Subjects = new ArrayList();
            ArrayList NumbersAndDates = new ArrayList();
            ArrayList Declarants = new ArrayList();
            //appeals = new ArrayList();
            // Объект для хранения общих данных обращения
            Appeal NewAppeal = new Appeal();
            NewAppeal.init();
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
                        if (ParseSubject(cellText, out cellParsedValues)) {
                            // Проверка завершилась успешно - заполняем массив субъектов
                            foreach (string str in cellParsedValues) {
                                Subjects.Add(str);
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add(new ErrorMessage("Субъект Российской Федерации: ", MessageType.Header, colIndex));
                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedValues) {
                                errors.Add(new ErrorMessage(str, MessageType.Text));
                            }
                        }
                        break;
                    // Содержание+
                    case 2:
                        if (ParseContent(cellText, out tmp)) {
                            NewAppeal.content = tmp;
                        } else {
                            result = false;
                            errors.Add(new ErrorMessage("Содержание: ", MessageType.Header, colIndex));
                            errors.Add(new ErrorMessage(tmp, MessageType.Text));
                        }
                        break;
                    // Заявитель+
                    case 3:
                        if (ParseDeclarant(cellText, out cellParsedValues)) {
                            // Проверка завершилась успешно - заполняем массив субъектов
                            foreach (string str in cellParsedValues) {
                                Declarants.Add(str);
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add(new ErrorMessage("Кем заявлено: ", MessageType.Header, colIndex));

                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedValues) {
                                errors.Add(new ErrorMessage(str,MessageType.Text));
                            }
                        }
                        break;
                    // Сведения о подтверждении
                    case 4:
                        if (ParseConfirmation(cellText, out tmp)) {
                            NewAppeal.confirmation = tmp;
                        } else {
                            result = false;
                            errors.Add(new ErrorMessage("Сведения о подтверждении: ", MessageType.Header, colIndex));
                            errors.Add(new ErrorMessage(tmp, MessageType.Text));
                        }
                        break;
                    // Приянтые меры
                    case 5:
                        if (ParseMeasures(cellText, out tmp)) {
                            NewAppeal.measures = tmp;
                        } else {
                            result = false;
                            errors.Add(new ErrorMessage("Приянтые меры: ", MessageType.Header, colIndex));
                            errors.Add(new ErrorMessage(tmp, MessageType.Text));
                        }
                        break;
                    // Номер и дата
                    case 6:
                        if (ParseNumberAndDate(cellText, out cellParsedValues)) {
                            // Проверка завершилась успешно - заполняем массив субъектов
                            foreach (Tuple<string, string> NumDate in cellParsedValues) {
                                NumbersAndDates.Add(NumDate);
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add(new ErrorMessage("Рег. номер и дата: ", MessageType.Header, colIndex));
                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedValues) {
                                errors.Add(new ErrorMessage(str, MessageType.Text));
                            }
                        }
                        break;
                    // 7 - Уровень выборов
                    // Партия
                    case 8:
                        if (ParseParty(cellText, out cellParsedValues)) {
                            foreach (string str in cellParsedValues) {
                                NewAppeal.multi.Add(new string[] { "tematika", str });
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add(new ErrorMessage("Партия: ", MessageType.Header, colIndex));
                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedValues) {
                                errors.Add(new ErrorMessage(str, MessageType.Text));
                            }
                        }

                        break;
                    // Тип заявителя
                    case 9:
                        if (ParseDeclarantType(cellText, out tmp)) {
                            NewAppeal.declarant_type = tmp;
                        } else {
                            result = false;
                            errors.Add(new ErrorMessage("Тип заявителя: ", MessageType.Header, colIndex));
                            errors.Add(new ErrorMessage(tmp, MessageType.Text));
                        }

                        break;
                    // Тематика
                    case 10:
                        if (ParseTheme(cellText, out cellParsedValues)) {
                            foreach(string str in cellParsedValues) {
                                NewAppeal.multi.Add(new string[] { "tematika", str });
                            }
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add(new ErrorMessage("Тематика: ", MessageType.Header, colIndex));

                            // Добавляем все сообщения об ошибках в errors.
                            foreach (string str in cellParsedValues) {
                                errors.Add(new ErrorMessage(str, MessageType.Text));
                            }
                        }
                        break;
                    // 11 - +
                    // Исполнитель
                    case 12:
                        if (ParseExecutor(cellText, out tmp)) {
                            NewAppeal.executor_id = tmp;
                        } else {
                            result = false;
                            errors.Add(new ErrorMessage("Исполнитель: ", MessageType.Header, colIndex));
                            errors.Add(new ErrorMessage(tmp, MessageType.Text));
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
            if (inputData == "") {
                resultData.Add("Не заполнено обязательное поле; ");
                return false;
            }

            // Удаляем лишние символы.
            inputData = PrepareRawData(inputData);
            // Разделяем текст на части. В Words будут записаны названия субъектов.
            char[] EntrySeparators = { ';' };
            string[] Words = inputData.Split(EntrySeparators, StringSplitOptions.RemoveEmptyEntries);
            char[] WordSeparators = { ' ', '-', '–' };

            // Ищем код каждого субъекта по его названию.
            foreach (string AppealSubjectName in Words) {
                string SubjCode = "";
                foreach (KeyValuePair<string, string> DirSubj in Subjects) {

                    foreach (string AppealSubjWord in AppealSubjectName.Split(WordSeparators, StringSplitOptions.RemoveEmptyEntries)) {
                        if (DirSubj.Key.Split(WordSeparators, StringSplitOptions.RemoveEmptyEntries).Contains(AppealSubjWord)) {
                            SubjCode = DirSubj.Value;
                        } else {
                            // Слово не найдено в названии субъекта из справочника
                            SubjCode = "";
                            break;
                        }
                    }
                    if (SubjCode != "") {
                        break;
                    }
                }
                if (SubjCode != "") {
                    // Субъект найден
                    resultData.Add(SubjCode);
                    break;
                } else {
                    resultData.Clear();
                    resultData.Add("Наименование субъекта РФ не найдено в справочнике \"Субъекты РФ\"; ");
                    return false;
                }
                //KeyValuePair<string, string> Subj = Subjects.First(s => s.Key == SubjectName.Trim());
                //resultData.Add(Subj.Value);
            }
            result = true;

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
        bool ParseDeclarant(string inputData, out ArrayList resultData) {
            resultData = new ArrayList();
            ArrayList AppealDeclarants = new ArrayList();
            
            // Проверяем, что поле заполнено
            if (inputData == "") {
                resultData.Add("Не заполнено обязательное поле; ");
                return false;
            }
            
            // Удаляем лишние символы.
            inputData = PrepareRawData(inputData);

            // Разделяем текст на части. В Words будут записаны заявители.
            char[] Separators = { ';' };
            string[] Words = inputData.Split(Separators, StringSplitOptions.RemoveEmptyEntries);
            
            // Проверяем формат.
            string pattern = @"\p{IsCyrillic}\.\p{IsCyrillic}\. \w+";
            string name, info;
            foreach (string str in Words) {
                name = "";
                info = "";
                if (str.IndexOf(",") >= 0) {
                    name = str.Substring(0, str.IndexOf(",")).Trim();
                    info = str.Substring(str.IndexOf(",") + 1).Trim();
                } else {
                    name = str;
                }
                if (Regex.IsMatch(name, pattern)) {
                    AppealDeclarants.Add(new string[] {name, info});
                } else {
                    // Неверный формат
                    resultData.Add( "Данные не соответствуют формату; " );
                    return false;
                }
            }

            // Ищем в справочнике id каждого заявителя по его имени.
            bool NewDeclarant;
            foreach (string[] declarant in AppealDeclarants) {
                NewDeclarant = true;
                foreach (string[] str in Declarants) {
                    if (str[1] == declarant[0]) {
                        // Заявитель найден в справочнике.
                        // str[0] - id заявителя.
                        resultData.Add(str[0]);
                        NewDeclarant = false;
                        break;
                    }
                }
                if (NewDeclarant) {
                    // Заявитель не был найден в справочнике - добавляем с фиктивным ID.
                    Declarants.Add(new string[] { "fake-" + DeclarantFakeID.ToString(), declarant[0] });
                    DeclarantFakeID++;
                }
            }
            return true;
        }
        bool ParseConfirmation(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                data = PrepareRawData(data);
                try {
                    KeyValuePair<string, string> Confirmation = Confirmations.First(s => s.Key == data);
                    parsed = Confirmation.Value;
                    result = true;
                } catch (System.InvalidOperationException) {
                    parsed = "Сведения не соответствуют классификатору \"Сведения о подтверждении\"; ";
                }
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
        bool ParseNumberAndDate(string inputData, out ArrayList resultData) {
            resultData = new ArrayList();
            ArrayList AppealDeclarants = new ArrayList();

            // Проверяем, что поле заполнено
            if (inputData == "") {
                resultData.Add("Не заполнено обязательное поле; ");
                return false;
            }

            // Удаляем лишние символы.
            inputData = inputData.Replace("№", "");
            inputData = PrepareRawData(inputData);

            // Разделяем текст на части. В Words будут записаны номера+даты.
            char[] Separators = { ';' };
            string[] Words = inputData.Split(Separators, StringSplitOptions.RemoveEmptyEntries);

            // Проверяем формат.
            string pattern = @"^\w+ от \d{1,2}\.\d{1,2}\.\d{4}$";
            string num, f_date;
            string trimmed;
            foreach (string str in Words) {
                trimmed = str.Trim();
                if (!Regex.IsMatch(trimmed, pattern)) {
                    resultData.Clear();
                    resultData.Add("Данные не соответствуют формату; ");
                    return false;
                }
                num = trimmed.Substring(0, trimmed.IndexOf(" от ")).Trim();
                f_date = trimmed.Substring(trimmed.IndexOf(" от ") + 1).Trim();
                resultData.Add(Tuple.Create(num, f_date));
            }
            return true;
        }
        bool ParseParty(string inputData, out ArrayList resultData) {
            bool result = true;
            resultData = new ArrayList();
            // Удаляем лишние символы.
            inputData = PrepareRawData(inputData);
            // Разделяем текст на части. В Words будут записаны названия субъектов.
            char[] Separatos = { ';' };
            string[] Words = inputData.Split(Separatos, StringSplitOptions.RemoveEmptyEntries);

            // Ищем код каждого субъекта по его названию.
            try {
                foreach (string str in Words) {
                    KeyValuePair<string, string> Party = Parties.First(s => s.Key == str.Trim());
                    resultData.Add(Party.Value);
                }
            } catch (System.InvalidOperationException) {
                result = false;
                resultData.Add("Тематика не соответствует классификатору \"Тематика\"; ");
            }

            return result;
        }
        bool ParseDeclarantType(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
            } else {
                data = PrepareRawData(data);
                try {
                    KeyValuePair<string, string> DecType = DecTypes.First(s => s.Key == data);
                    parsed = DecType.Value;
                    result = true;
                } catch (System.InvalidOperationException) {
                    parsed = "Тип заявителя не соответствуют классификатору \"Тип заявителей\"; ";
                }
            }
            return result;

        }
        bool ParseExecutor(string data, out string parsed) {
            bool result = false;
            if (data == "") {
                parsed = "Не заполнено обязательное поле; ";
                return false;
            }
            data = PrepareRawData(data);
            string ShortPattern = @"^\w+$";
            string FullPattern = @"^\w+ \p{IsCyrillic}\.\p{IsCyrillic}\.$";

            if (Regex.IsMatch(data, FullPattern) || Regex.IsMatch(data, ShortPattern)) {
                // Формат подходит, проверяем наличие в справочнике
                try {
                    KeyValuePair<string, string> Executor = Executors.First(s => s.Key == data);
                    parsed = Executor.Value;
                    result = true;
                } catch (System.InvalidOperationException) {
                    parsed = "Исполнитель не найден в справочнике \"Список исполнителей\"; ";
                    return false;
                }
            } else {
                parsed = "Данные не соответствуют формату; ";
            }

            return result;
        }
        bool ParseTheme(string inputData, out ArrayList resultData) {
            bool result = true;
            resultData = new ArrayList();
            if (inputData == "") {
                resultData.Add("Не заполнено обязательное поле; ");
                result = false;
            }
            // Удаляем лишние символы.
            inputData = PrepareRawData(inputData);
            // Разделяем текст на части. В Words будут записаны названия субъектов.
            char[] Separatos = { ';' };
            string[] Words = inputData.Split(Separatos, StringSplitOptions.RemoveEmptyEntries);

            // Ищем код каждого субъекта по его названию.
            try {
                foreach (string str in Words) {
                    KeyValuePair<string, string> Theme = Themes.First(s => s.Key == str.Trim());
                    resultData.Add(Theme.Value);
                }
            } catch (System.InvalidOperationException) {
                result = false;
                resultData.Add("Тематика не соответствует классификатору \"Тематика\"; ");
            }

            return result;
        }
        string PrepareRawData(String data) {
            // Удаляет лишние символы
            string result = Regex.Replace(data, @"\s+", " ");
            return result.Trim();
        }
        bool LoadDictionaries() {
            string[] args = Environment.GetCommandLineArgs();
            string DBName = "RA00C000";
            string DBUser = "voshod";
            string DBPass = "voshod";
            if (args.Length > 3) {
                DBName = args[1];
                DBUser = args[2];
                DBPass = args[3];
            }
            string oradb = "Data Source=" + DBName + ";User Id=" + DBUser + ";Password=" + DBPass + ";";
            OracleConnection conn = null;
            try {
                conn = new OracleConnection(oradb);
                conn.Open();
            } catch (Exception e) {
                MessageBox.Show("Не удалось подключиться к базе данных.\n" + e.Message, "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
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
            }

            // Подтверждения.
            cmd.CommandText = "select content, TO_CHAR(id) from akriko.cls_podtv order by id";
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
                    Confirmations.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                }
            }

            // Тематики.
            cmd.CommandText = "select REPLACE(numb, '.', ''), TO_CHAR(id) from akriko.cls_tematika order by numb";
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
                    Themes.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                }
            }

            // Партии.
            for (int i = 1; i < 75; i++) {
                Parties.Add(i.ToString(), "1001000" + (100 + i).ToString());
            }

            // Типы заявителей.
            cmd.CommandText = "select REPLACE(numb, '.', ''), TO_CHAR(id) from akriko.cls_zayaviteli order by numb";
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
                    DecTypes.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                }
            }

            // Исполнители.
            cmd.CommandText = "select l_name, TO_CHAR(id) from akriko.cat_executors order by l_name";
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
                    Executors.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                }
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

    class ErrorMessage {
        public readonly string Message;
        public readonly MessageType Type;
        public readonly int Column;

        public ErrorMessage(string msg, MessageType t, int c = -1) {
            this.Message = msg;
            this.Type = t;
            this.Column = c;
        }
    }
}
