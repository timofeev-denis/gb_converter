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
        const string PROGRAM_DIR = @"e:\akriko_converter\!_dev\";

        Dictionary<string, string> Subjects = new Dictionary<string, string>();


        static void _t( String str ) {
            System.Diagnostics.Trace.WriteLine( str );
        }
        static void Main(string[] args) {
            Program __instance = new Program();
            __instance.Convert();
        }
        void Convert() {
            string oradb = "Data Source=RA00C000;User Id=voshod;Password=voshod;";
            OracleConnection conn = new OracleConnection(oradb);
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "select namate, subjcod from ate_history where prsubj='1' and datedel is null";
            cmd.CommandType = CommandType.Text;
            OracleDataReader dr = null;
            try {
                dr = cmd.ExecuteReader();
            } catch (Oracle.DataAccess.Client.OracleException e) {
                _t(e.Message.ToString());
                cmd.Dispose();
                conn.Dispose();
                return;
            }

            while (dr.Read()) {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1)) {
                    Subjects.Add(dr.GetString(0).Trim(), dr.GetString(1).Trim());
                }
            }
            dr.Dispose();
            cmd.Dispose();
            conn.Dispose();



            /*
            if (!this.LoadDictionaries()) {
                return;
            }
             * */

            ArrayList RowErrors = new ArrayList();
            DocX document = DocX.Load(PROGRAM_DIR + "фрагмент зеленой книги_1 (1).docx");
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
                if (CheckAppeal(appealsTable.Rows[rowIndex], out RowErrors)) {
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
        bool CheckAppeal(Row row, out ArrayList errors) {
            bool result = true;
            string cellText = "";
            errors = new ArrayList();
            ArrayList cellParsedText;
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
                switch (colIndex) {
                    case 1:
                        if (ParseSubject(cellText, out cellParsedText)) {
                            // Проверка завершилась успешно
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
            return true;
        }
    }

}
