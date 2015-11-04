using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using Oracle.DataAccess.Client;
using System.Data;
using System.Collections;
using System.Text.RegularExpressions;

namespace DocxText {
    class Program {

        const int RESULT_COLUMN = 13;
        //const string PROGRAM_DIR = @"e:\akriko_converter\!_dev\";
        const string PROGRAM_DIR = @"f:\@job\@Voskhod\AKRIKO\";
        //const string SOURCE_FILE = @"фрагмент зеленой книги_1 (1).docx";
        const string SOURCE_FILE = "gb1.docx";

        Dictionary<string, string> Subjects = new Dictionary<string, string>();


        static void _t( String str ) {
            System.Diagnostics.Trace.WriteLine( str );
        }
        static void Main(string[] args) {
            Program __instance = new Program();
            __instance.Convert();
        }
        void Convert() {
            // Загружаем справочник субъектов РФ из БД
            if (!this.LoadDictionaries()) {
                return;
            }

            ArrayList RowErrors = new ArrayList();
            ArrayList SimpleAppeals = new ArrayList();
            
            DocX document = DocX.Load(PROGRAM_DIR + SOURCE_FILE);
            Table appealsTable = document.Tables[0];

            // Добавляем колонку с информацией о выявленных несоответствиях по обращению (результат проверки).
            appealsTable.InsertColumn();
            appealsTable.Rows[0].Cells[RESULT_COLUMN].VerticalAlignment = VerticalAlignment.Top;
            Paragraph p = appealsTable.Rows[0].Cells[RESULT_COLUMN].Paragraphs[0].Append("Выявленные несоответствия данных (для конвертации)");
            p.Alignment = Alignment.center;
            p.Bold();

            // Анализируем все строки таблицы.
            for (int rowIndex = 1; rowIndex < appealsTable.Rows.Count; rowIndex++) {
                // checkResult = CheckAppeal(appealsTable.Rows[rowIndex]);
                if (CheckAppeal(appealsTable.Rows[rowIndex], out SimpleAppeals, out RowErrors)) {
                    // Проверка обращения прошла успешно.

                } else {
                    // Проверка обращения завершилась с ошибками.
                    // Записываем результат проверки в правую колонку.
                    foreach (string str in RowErrors) {
                        appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs[0].Append(str);
                    }
                    //p.FontSize(10);
                }
            }

            document.SaveAs(PROGRAM_DIR + "Ошибки конвертации " + (DateTime.Now).ToString("yyyy-MM-dd-HH-mm-ss") + ".docx");
            // System.Diagnostics.Trace.WriteLine((DateTime.Now).ToString( "yyyy-MM-dd-H-mm-ss" ));
        }
        /// <summary>
        /// <para>Проверяет данные заявки</para>
        /// <para>Возвращает true если данные разобраны без ошибок и false в противном случае.</para>
        /// </summary>
        /// <param name="row">Строка таблицы.</param>
        /// <param name="parsedData">out ArayList с данными либо с ошибками, если разбор завершился неудаче.</param>
        bool CheckAppeal(Row row, out ArrayList appeals, out ArrayList errors) {
            bool result = true;
            string cellText = "";
            errors = new ArrayList();
            ArrayList cellParsedText;
            ArrayList Subjects = new ArrayList();
            ArrayList NumbersAndDates = new ArrayList();
            appeals = new ArrayList();
            // Пропускаем первую и последнюю колонку.
            for (int colIndex = 1; colIndex < row.Cells.Count - 1; colIndex++) {
                Cell c = row.Cells[colIndex];
                // Собираем текст ячейки из всех параграфов в одну переменную.
                for (int i = 0; i < c.Paragraphs.Count; i++) {
                    if (c.Paragraphs[i].Text.Trim() != "") {
                        cellText += c.Paragraphs[i].Text + " ";
                    }
                }
                // Запускаем разбор текста из ячейки.
                Appeal NewAppeal = new Appeal();
                switch (colIndex) {
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
                            foreach(string str in cellParsedText ) {
                                errors[errors.Count - 1] += str + "; ";
                            }
                        }
                        break;
                }
            }

            // Создаём "простые" обращения
            foreach (string SubjCode in Subjects) {
                foreach (string NumDate in NumbersAndDates) {
                    appeals.Add(new Appeal());
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
                //result = true;
            } catch (System.InvalidOperationException) {
                resultData.Clear();
                resultData.Add( "Наименование субъекта РФ не найдено в справочнике \"Субъекты РФ\"" );
            }

            return result;
        }
        static bool ParseContent(string content) {
            bool result = false;
            return result;
        }
        static bool ParseDeclarant(string declarant) {
            bool result = false;
            return result;
        }
        static bool ParseConfirmation(string confirmation) {
            bool result = false;
            return result;
        }
        static bool ParseMeasures(string measures) {
            bool result = false;
            return result;
        }
        static bool ParseNumberAndDate(string appealNumber, string appealDate) {
            bool result = false;
            //string sPattern = @"^\d{3}-\d{3}-\d{4}$";
            return result;
        }
        static bool ParseParty(string party) {
            bool result = false;
            return result;
        }
        static bool ParseDeclarantType(string declarantType) {
            bool result = false;
            return result;
        }
        static bool ParseExecutor(string executor) {
            bool result = false;
            return result;
        }
        static bool ParseTheme(string theme) {
            bool result = false;
            return result;
        }
        static string PrepareRawData(String data) {
            string result = Regex.Replace(data, @"\s+", " ");
            return result;
        }
        bool LoadDictionaries() {
            
            // UNCOMMENT!

            /*
            string oradb = "Data Source=RA00C000;User Id=voshod;Password=voshod;";
            OracleConnection conn = new OracleConnection(oradb);
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "select namate, subjcode from ate_history where prsubj='1' and datedel is null";
            cmd.CommandType = CommandType.Text;
            OracleDataReader dr = null;
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                conn.Dispose();
                return false;
            }

            while (dr.Read()) {
                Subjects.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                //subjects.FirstOrDefault()
            }
            dr.Dispose();
            cmd.Dispose();
            conn.Dispose();
            */

            // DB emulation :)
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

            return true;
        }
        /// <summary>
        /// Класс для работы с обращениями АКРИКО
        /// </summary>
        class Akriko {
            public enum TableName { appeal, cat_executors, cat_declarants }

            int AddAppeal(Appeal appeal) {
                // content & content_cik - ?
                // {7} - created
                string Query = "INSERT INTO akriko.appeal " +
                    "(id, numb, f_date, content, subjcode, parent_id, meri, created, content_cik, ispolnitel_cik_id) " +
                    String.Format(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')",
                    GetID(TableName.appeal), appeal.numb, appeal.f_date, appeal.content,
                    appeal.subjcode, appeal.parent_id, appeal.measures, appeal.created,
                    appeal.content_cik, appeal.executor_id);

                return 0;
            }
            public string GetID(TableName t) {
                /// 1. Считать значение последовательности.
                /// 2. Перевести в строку.
                /// 3. Объединить с "100".
                // Соединяемся с БД или используем общее соединение.
                // ...
                // Считываем значение последовательности.
                long sequence;
                switch (t) {
                    case TableName.appeal:
                        break;
                    case TableName.cat_declarants:
                        break;
                    case TableName.cat_executors:
                        break;
                }
                // Для отладки.
                Random R = new Random();

                sequence = R.Next(1000, 2000);
                //new Random().Next(1000, 2000);
                return "100" + sequence.ToString();
            }
        }
    }
    public class Appeal {
        public string id;
        public string numb;
        public string f_date;
        public string content;
        public string subjcode;
        public string parent_id;
        public string measures;
        public string created;
        public string content_cik;
        public string executor_id;
    }


}
