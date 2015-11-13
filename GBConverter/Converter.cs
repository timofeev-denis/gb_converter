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
        List<Declarant> Declarants = new System.Collections.Generic.List<Declarant>();
        long DeclarantFakeID = 1;
        private ArrayList SimpleAppeals = new ArrayList();
        private DateTime ConvertDate;

        static void _t(String str) {
            System.Diagnostics.Trace.WriteLine(str);
        }

        public bool CheckFile(string fileName, ProgressBar progressBar = null) {
            bool result = false;
            if (fileName == "") {
                throw new ArgumentException("Не указано имя файла для конвертации");
            }
            // Загружаем справочники из БД
            if (!this.LoadDictionaries()) {
                return false;
            }

            ArrayList RowErrors = new ArrayList();
            this.SimpleAppeals.Clear();
            DocX document;
            try {
                // Открываем файл с "ЗЕлёной книгой".
                document = DocX.Load(fileName);
            } catch (System.IO.IOException e) {
                MessageBox.Show(String.Format("Не удалось открыть файл {0}.\nВозможно он используется другой программой.\n\nРабота программы прекращена.", fileName),
                    "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
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
                if (!CheckAppeal(appealsTable.Rows[rowIndex], SimpleAppeals, out RowErrors)) {
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
                            /*
                            if (msg.Column > 0) {
                                Color violetColor = Color.FromArgb(204, 0, 153);
                                foreach (Paragraph ColorParagraph in appealsTable.Rows[rowIndex].Cells[msg.Column].Paragraphs) {
                                    if (ColorParagraph.Text.Trim() != "") {
                                        ColorParagraph.Color(violetColor);
                                    }
                                }
                            }
                            */
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
                if (progressBar != null) {
                    double percent = (double)rowIndex / (appealsTable.Rows.Count - 1) * 100;
                    progressBar.Value = System.Convert.ToInt32(percent);
                }
            }

            if (Step2) {
                // Несоответствий не выявлено.
                result = true;
            } else {
                string filePath = Path.GetDirectoryName(fileName) + @"\Ошибки конвертации " + (DateTime.Now).ToString("yyyy-MM-dd-HH-mm-ss") + ".docx";
                document.SaveAs(filePath);
                System.Diagnostics.Process.Start(filePath);
                result = false;
            }
            document.Dispose();
            return result;
            // System.Diagnostics.Trace.WriteLine((DateTime.Now).ToString( "yyyy-MM-dd-H-mm-ss" ));
        }
        /// <summary>
        /// <para>Проверяет данные заявки</para>
        /// <para>Возвращает true если данные разобраны без ошибок и false в противном случае.</para>
        /// </summary>
        /// <param name="row">Строка таблицы.</param>
        /// <param name="parsedData">out ArayList с данными либо с ошибками, если разбор завершился неудаче.</param>
        bool CheckAppeal(Row row, ArrayList appeals, out ArrayList errors) {
            errors = new ArrayList();
            bool result = true;
            string cellText = "";
            string tmp;
            string DeclarantType = "";
            string DeclarantParty = "";
            ArrayList cellParsedValues;
            ArrayList Subjects = new ArrayList();
            ArrayList NumbersAndDates = new ArrayList();
            // ФИО и info заявителей из колонки "Кем заявлено".
            ArrayList AppealDeclarants = new ArrayList();
            // Объект для хранения общих данных обращения/обращений
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
                            // Проверка завершилась успешно - заполняем массив заявителей
                            foreach (string[] str in cellParsedValues) {
                                //Declarants.Add(str);
                                AppealDeclarants.Add(str);
                                //NewAppeal.multi.Add(new string[] { "declarant", str });
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
                        if (ParseParty(cellText, out tmp)) {
                            DeclarantParty = tmp;
                            //NewAppeal.multi.Add(new string[] { "tematika", tmp });
                        } else {
                            // Проверка завершилась с ошибкой.
                            result = false;
                            errors.Add(new ErrorMessage("Партия: ", MessageType.Header, colIndex));
                            // Добавляем все сообщения об ошибках в errors.
                            errors.Add(new ErrorMessage(tmp, MessageType.Text));
                        }

                        break;
                    // Тип заявителя
                    case 9:
                        if (ParseDeclarantType(cellText, out tmp)) {
                            //NewAppeal.declarant_type = tmp;
                            DeclarantType = tmp;
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

            // Проверяем заявителей
            // AppealDeclarants - ArrayList<string[fio, info]>
            // DeclarantParty - id партии всех заявителей
            // DeclarantType - id типа заявителя всех заявителей
            // Пройтись по AppealDeclarants, сформировать каждого заявителя, поискать его в справочнике (если надо - создать)
            // В NewAppeal.multi добавить информацию о заявителях с реальными id
            bool NewDeclarant;
            string DeclarantID = "";
            foreach (string[] fio_info in AppealDeclarants) {
                NewDeclarant = true;
                Declarant d = new Declarant(DeclarantType, DeclarantParty, fio_info[1]);
                d.SetFIO(fio_info[0]);
                // Ищем в справочнике.
                foreach (Declarant dd in Declarants) {
                    if (d.Equals(dd)) {
                        NewDeclarant = false;
                        DeclarantID = dd.GetID();
                        break;
                    }
                }
                if (NewDeclarant) {
                    // Заводим нового заявителя.
                    DeclarantID = d.SaveToDB();
                    Declarants.Add(d);
                }
                // Доавляем заявителя к обращению.
                NewDeclarant = true;
                foreach (string[] str in NewAppeal.multi) {
                    if (str[0] == "declarant" && str[1] == DeclarantID) {
                        NewDeclarant = false;
                        break;
                    }
                }

                if (NewDeclarant) {
                    NewAppeal.multi.Add(new string[] { "declarant", DeclarantID });
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
        bool CreateAppeal(Appeal newAppeal) {
            OracleConnection conn = DB.GetConnection();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            OracleDataReader dr = null;
            string NewAppealID = "";
            //
            
            cmd.CommandText = "select CONCAT('1" + newAppeal.subjcode + "', LPAD(AKRIKO.seq_appeal.NEXTVAL,7,'0')) from dual";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                DB.CloseConnection();
                return false;
            }
            dr.Read();
            if (dr.IsDBNull(0)) {
                DB.CloseConnection();
                return false;
            } else {
                NewAppealID = dr.GetString(0);
            }
            
            //
            cmd.CommandType = CommandType.Text;
            String query = "insert into AKRIKO.APPEAL " +
                "(id, numb, f_date, hod_ispoln, is_control, is_repeat, podtv, subjcode, is_sud, is_collective, replicate_need, created, unread, meri_cik, links, sud_tematika, content_cik, ispolnitel_cik_id, del, only_sud)" +
                " VALUES (:newappealid, :numb, TO_DATE(:f_date, 'DD.MM.YYYY'), :hod_ispoln, :is_control, :is_repeat, :podtv, :subjcode, :is_sud, :is_collective, :replicate_need, :created, :unread, :meri_cik, :links, :sud_tematika, :content_cik, :ispolnitel_cik_id, :del, :only_sud)";
            
            OracleCommand command = new OracleCommand(query, conn);
            command.Parameters.Add(":newappealid", NewAppealID);
            command.Parameters.Add(":numb", newAppeal.numb);
            command.Parameters.Add(":f_date", newAppeal.f_date);
            //command.Parameters.Add(":f_date", this.ConvertDate);
            command.Parameters.Add(":hod_ispoln", "0");
            command.Parameters.Add(":is_control", "0");
            command.Parameters.Add(":is_repeat", "0");
            command.Parameters.Add(":podtv", newAppeal.confirmation);
            command.Parameters.Add(":subjcode", Int32.Parse(newAppeal.subjcode));
            command.Parameters.Add(":is_sud", "0");
            command.Parameters.Add(":is_collective", "0");
            command.Parameters.Add(":replicate_need", "0");
            command.Parameters.Add(":created", this.ConvertDate);
            command.Parameters.Add(":unread", "1");
            command.Parameters.Add(":meri_cik", newAppeal.measures);
            command.Parameters.Add(":links", "0");
            command.Parameters.Add(":sud_tematika", "0");
            command.Parameters.Add(":content_cik", newAppeal.content);
            command.Parameters.Add(":ispolnitel_cik_id", newAppeal.executor_id);
            command.Parameters.Add(":del", "0");
            command.Parameters.Add(":only_sud", "0");
            
            command.ExecuteNonQuery();

            foreach (string[] str in newAppeal.multi) {
                command.CommandText = "insert into akriko.appeal_multi (appeal_id,col_name,content,key) values(" + NewAppealID + ",'" + str[0] + "','" + str[1] + "',0)";
                command.ExecuteNonQuery();
            }
            command.Dispose();

            return true;
        }
        public bool Convert(ProgressBar progressBar = null) {
            if (progressBar != null) {
                progressBar.Value = 0;
            }
            bool result = true;
            int AppealIndex = 1;
            // Запоминаем текущую дату и время, чтобы установить их всем создаваемым обращениям.
            this.ConvertDate = new DateTime();
            foreach (Appeal NewAppeal in this.SimpleAppeals) {
                if (!CreateAppeal(NewAppeal)) {
                    result = false;
                }
                if (progressBar != null) {
                    double percent = (double)AppealIndex / this.SimpleAppeals.Count * 100;
                    progressBar.Value = System.Convert.ToInt32(percent);
                }
                AppealIndex++;
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
            string FullPattern = @"\p{IsCyrillic}\.\p{IsCyrillic}\. \w+";
            string ShortPattern = @"\p{IsCyrillic}\. \w+";
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
                if (name.ToLower() == "коллективное обращение" || Regex.IsMatch(name, FullPattern) || Regex.IsMatch(name, ShortPattern)) {
                    resultData.Add(new string[] { name, info });
                } else {
                    // Неверный формат
                    resultData.Add( "Данные не соответствуют формату; " );
                    return false;
                }
            }
            /*
            // Ищем в справочнике id каждого заявителя по его имени.
            bool NewDeclarant;
            foreach (string[] declarant in AppealDeclarants) {
                NewDeclarant = true;
                foreach (Declarant d in Declarants) {
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
                    // Добавляем в результирующий список заявителей
                    resultData.Add("fake-" + DeclarantFakeID.ToString());
                    DeclarantFakeID++;
                }
            }
            */
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
            inputData = inputData.ToLower();

            // Разделяем текст на части. В Words будут записаны номера+даты.
            char[] Separators = { ';' };
            string[] Words = inputData.Split(Separators, StringSplitOptions.RemoveEmptyEntries);

            // Проверяем формат.
            string pattern = @"^[а-яА-Я0-9\./-]+ от \d{1,2}\.\d{1,2}\.\d{4}$";
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
                f_date = trimmed.Substring(trimmed.IndexOf(" от ") + 4).Trim();
                resultData.Add(Tuple.Create(num, f_date));
            }
            return true;
        }
        bool ParseParty(string inputData, out string resultData) {
            bool result = true;
            resultData = "";
            // Удаляем лишние символы.
            inputData = PrepareRawData(inputData);

            // Ищем код партии по её названию.
            try {
                KeyValuePair<string, string> Party = Parties.First(s => s.Key == inputData.Trim());
                resultData = Party.Value;
            } catch (System.InvalidOperationException) {
                resultData = "Значение не найдено в справочнике \"Общественные объединения\"; ";
                return false;
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
                    if(data.IndexOf(' ' ) >= 0) {
                        data = data.Substring(0, data.IndexOf(' '));
                    }
                    KeyValuePair<string, string> Executor = Executors.First(s => s.Key == data.ToLower());
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
            OracleConnection conn = DB.GetConnection();

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
                DB.CloseConnection();
                return false;
            }
            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    try {
                        Subjects.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                    } catch (Exception e) {
                        MessageBox.Show(dr.GetString(0).Trim(), "Найден дубликат субъекта");
                    }
                }
            }

            // Подтверждения.
            cmd.CommandText = "select content, TO_CHAR(id) from akriko.cls_podtv order by id";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                DB.CloseConnection();
                return false;
            }
            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    try {
                        Confirmations.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                    } catch (Exception e) {
                        MessageBox.Show(dr.GetString(0).Trim(), "Найден дубликат подтверждения");
                    }

                }
            }

            // Тематики.
            cmd.CommandText = "select REPLACE(numb, '.', ''), TO_CHAR(id) from akriko.cls_tematika order by numb";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                DB.CloseConnection();
                return false;
            }
            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    try { 
                        Themes.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                    } catch (Exception e) {
                        MessageBox.Show(dr.GetString(0).Trim(), "Найден дубликат тематики");
                    }
                }
            }

            // Партии.
            Parties.Add("1", "199199934542");
            Parties.Add("2", "199199934548");
            Parties.Add("3", "199199934551");
            Parties.Add("4", "199199934589");
            Parties.Add("5", "199199934555");
            Parties.Add("6", "199199934553");
            Parties.Add("7", "100100023740549");
            Parties.Add("8", "100100035781415");
            Parties.Add("9", "100100035780274");
            Parties.Add("10", "100100035780277");
            Parties.Add("11", "100100035781337");
            Parties.Add("12", "100100035781340");
            Parties.Add("13", "100100035781343");
            Parties.Add("14", "100100035781346");
            Parties.Add("15", "100100035781412");
            Parties.Add("16", "100100035781409");
            Parties.Add("17", "100100035784336");
            Parties.Add("18", "100100035784351");
            Parties.Add("19", "100100035784345");
            Parties.Add("20", "100100035784718");
            Parties.Add("21", "100100035786299");
            Parties.Add("22", "100100035786293");
            Parties.Add("23", "100100035786227");
            Parties.Add("24", "100100035786290");
            Parties.Add("25", "100100035786296");
            Parties.Add("26", "100100035786308");
            Parties.Add("27", "100100035786317");
            Parties.Add("28", "100100035786311");
            Parties.Add("29", "100100035786314");
            Parties.Add("30", "100100039749981");
            Parties.Add("31", "100100035786305");
            Parties.Add("32", "100100039750663");
            Parties.Add("33", "100100035790815");
            Parties.Add("34", "100100039749991");
            Parties.Add("35", "100100047868079");
            Parties.Add("36", "100100048439721");
            Parties.Add("37", "100100049142028");
            Parties.Add("38", "100100048439595");
            Parties.Add("39", "100100035790826");
            Parties.Add("40", "100100042277558");
            Parties.Add("41", "100100048439556");
            Parties.Add("42", "100100047878147");
            Parties.Add("43", "100100048439599");
            Parties.Add("44", "100100047868051");
            Parties.Add("45", "100100047878141");
            Parties.Add("46", "100100048439836");
            Parties.Add("47", "100100047878144");
            Parties.Add("48", "100100048439783");
            Parties.Add("49", "100100049649307");
            Parties.Add("50", "100100042277551");
            Parties.Add("51", "100100048442301");
            Parties.Add("52", "100100048442433");
            Parties.Add("53", "100100049141914");
            Parties.Add("54", "100100049282480");
            Parties.Add("55", "100100047878150");
            Parties.Add("56", "100100049649313");
            Parties.Add("57", "100100049100931");
            Parties.Add("58", "100100049649304");
            Parties.Add("59", "100100049141955");
            Parties.Add("60", "100100049649316");
            Parties.Add("61", "100100049282369");
            Parties.Add("62", "100100049282372");
            Parties.Add("63", "100100049691927");
            Parties.Add("64", "100100049735668");
            Parties.Add("65", "100100049282484");
            Parties.Add("66", "100100051160839");
            Parties.Add("67", "100100051161302");
            Parties.Add("68", "100100051161305");
            Parties.Add("69", "100100051348825");
            Parties.Add("70", "100100052561221");
            Parties.Add("71", "100100052080030");
            Parties.Add("72", "100100052561217");
            Parties.Add("73", "100100053307143");
            Parties.Add("74", "100100054451488");

            // Типы заявителей.
            cmd.CommandText = "select REPLACE(numb, '.', ''), TO_CHAR(id) from akriko.cls_zayaviteli order by numb";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                DB.CloseConnection();
                return false;
            }
            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    try { 
                        DecTypes.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                    } catch (Exception e) {
                        MessageBox.Show(dr.GetString(0).Trim(), "Найден дубликат типа заявителя");
                    }

                }
            }

            // Исполнители.
            cmd.CommandText = "select lower(l_name), TO_CHAR(id) from akriko.cat_executors where TO_CHAR(id) like '100%' order by l_name";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                DB.CloseConnection();
                return false;
            }
            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    try { 
                        Executors.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                        //File.AppendAllText(@".\log.txt", "executor - " + dr.GetString(0).Trim() + "\r\n");
                    } catch (Exception e) {
                        MessageBox.Show(dr.GetString(0).Trim(), "Найден дубликат исполнителя");
                    }

                }
            }
             
            // Заявители
            cmd.CommandText = "select SUBSTR(TRIM(f_name), 0, 1), NVL(SUBSTR(TRIM(m_name), 0, 1), ''), TRIM(l_name), type, NVL(TO_CHAR(party), ''), NVL(TRIM(info), ''), TO_CHAR(id) from akriko.cat_declarants where id = 1000000101";
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                DB.CloseConnection();
                return false;
            }
            while (dr.Read()) {
                try {
                    string fname = "";
                    string mname = "";
                    string lname = "";
                    string type = "";
                    string party = "";
                    string info = "";
                    string id = "";
                    if (!dr.IsDBNull(0)) {
                        fname = dr.GetString(0);
                    }
                    if (!dr.IsDBNull(1)) {
                        mname = dr.GetString(1);
                    }
                    if (!dr.IsDBNull(2)) {
                        lname = dr.GetString(2);
                    }
                    if (!dr.IsDBNull(3)) {
                        type = dr.GetValue(3).ToString();
                    }
                    if (!dr.IsDBNull(4)) {
                        party = dr.GetString(4);
                    }
                    if (!dr.IsDBNull(5)) {
                        info = dr.GetString(5);
                    }
                    if (!dr.IsDBNull(6)) {
                        id = dr.GetString(6);
                    }

                    Declarants.Add(new Declarant(fname, mname, lname, type, party, info, id));
                    //dr.GetString(0).Trim(), dr.GetString(1).Trim()
                } catch (Exception e) {
                    MessageBox.Show("При чтении из справочника заявителей возникла ошибка", "Конвертер Зелёной книги", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }

            /*
            Executors.Add("Алёшкин", "1");
            Executors.Add("Аракелян", "2");
            Executors.Add("Артамошкин", "3");
            Executors.Add("Белых", "4");
            Executors.Add("Бородулина", "5");
            Executors.Add("Воронин", "6");
            Executors.Add("Ермаков", "7");
            Executors.Add("Котомкин", "8");
            Executors.Add("Кубелун", "9");
            Executors.Add("Мешков", "10");
            Executors.Add("Неронов", "11");
            Executors.Add("Орловская", "12");
            Executors.Add("Пеетухов", "13");
            Executors.Add("Петурова", "14");
            Executors.Add("Петухоов", "15");
            Executors.Add("Попов", "16");
            Executors.Add("Симонова", "17");
            Executors.Add("соломонидина", "18");
            Executors.Add("Сомов", "19");
            Executors.Add("Соомов", "20");
            Executors.Add("Стоноженко", "21");
            Executors.Add("Токмачев", "22");
            Executors.Add("Тюняеева", "23");
            Executors.Add("Цветкова", "24");
            Executors.Add("Цветкоова", "25");
            Executors.Add("Чувина", "26");
            Executors.Add("Шеншин", "27");
            Executors.Add("артамошкин", "28");
            Executors.Add("афанасова", "29");
            Executors.Add("бабак", "30");
            Executors.Add("булгаков", "31");
            Executors.Add("егоров", "32");
            Executors.Add("захаров", "33");
            Executors.Add("копосов", "34");
            Executors.Add("кузяков", "35");
            Executors.Add("луценко", "36");
            Executors.Add("лученко", "37");
            Executors.Add("осадчая", "38");
            Executors.Add("поспеловская", "39");
            Executors.Add("пугачева", "40");
            Executors.Add("романова", "41");
            Executors.Add("стоноженко", "42");
            Executors.Add("тюняева", "43");
            Executors.Add("федотова", "44");
            Executors.Add("фуфаева", "45");
            Executors.Add("черкашина", "46");
            Executors.Add("якшина", "47");
            */
            dr.Dispose();
            cmd.Dispose();
            //DB.CloseConnection();


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
