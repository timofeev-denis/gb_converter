using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using Oracle.DataAccess.Client;
using System.Data;
using System.Collections;

namespace DocxText {
    class Program {

        const int RESULT_COLUMN = 13;
        const string PROGRAM_DIR = @"e:\akriko_converter\!_dev\";
        static void _t( String str ) {
            System.Diagnostics.Trace.WriteLine( str );
        }
        static void Main(string[] args) {            
            DocX document = DocX.Load(PROGRAM_DIR + "фрагмент зеленой книги_1 (1).docx");
            Table appealsTable = document.Tables[0];

            //string sPattern = @"^\d{3}-\d{3}-\d{4}$";
            string checkResult = "";

            // Добавляем ячейку с информацией о выявленных несоответствиях по обращению (результат проверки).
            appealsTable.InsertColumn();
            appealsTable.Rows[0].Cells[RESULT_COLUMN].VerticalAlignment = VerticalAlignment.Top;

            // Анализируем все строки таблицы.
            for (int rowIndex = 0; rowIndex < appealsTable.Rows.Count; rowIndex++) {
                if (rowIndex == 0) {
                    // Заголовок таблицы.
                    checkResult = "Выявленные несоответствия данных (для конвертации)";
                } else {
                    // Проверка данных.
                    //checkResult = CheckAppeal(appealsTable.Rows[rowIndex]);
                    
                }
                // Записываем результат проверки в правую колонку.
                if (checkResult != "") {
                    Paragraph p = appealsTable.Rows[rowIndex].Cells[RESULT_COLUMN].Paragraphs[0].Append(checkResult);
                    p.FontSize(10);
                    if (rowIndex == 0) {
                        p.Alignment = Alignment.center;
                        p.Bold();
                    }
                }
            }

            document.SaveAs(PROGRAM_DIR + "Ошибки конвертации " + (DateTime.Now).ToString( "yyyy-MM-dd-HH-mm-ss" ) + ".docx");
            // System.Diagnostics.Trace.WriteLine((DateTime.Now).ToString( "yyyy-MM-dd-H-mm-ss" ));
        }
        /// <summary>
        /// <para>Проверяет данные заявки</para>
        /// <para>Возвращает true если данные разобраны без ошибок и false в противном случае</para>
        /// </summary>
        /// <param name="row">Строка таблицы</param>
        /// <param name="parsedData">out ArayList с данными либо с ошибками, если разбор завершился неудаче</param>
        /// <param name="demoParam">DEMO</param>
        static bool CheckAppeal(Row row, out ArrayList parsedData) {
            bool result = true;
            string cellText = "";
            parsedData = new ArrayList();
            // Пропускаем первую и последнюю колонку.
            for (int colIndex = 1; colIndex < row.Cells.Count - 1; colIndex++) {
                Cell c = row.Cells[colIndex];
                for (int i = 0; i < c.Paragraphs.Count; i++) {
                    if (c.Paragraphs[i].Text.Trim() != "") {
                        cellText += c.Paragraphs[i].Text + " ";
                    }
                }

                switch (colIndex) {
                    case 1:
                        System.Diagnostics.Trace.Write(cellText);
                        //ValidateSubjcode(ref cellText);
                        System.Diagnostics.Trace.WriteLine(" -> " + cellText);
                        break;
                }
                //result += " | ";
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
        static bool ParseSubject(string inputData, out string[] resultData) {
            bool result = false;
            resultData = new string[10];
            ArrayList l = new ArrayList();
            l.Add("A");
            l.Add("B");
            System.Diagnostics.Trace.WriteLine(l.Count);
            return result;
        }
        static bool ParseExecutor(string executor) {
            bool result = false;
            return result;
        }
        static void OracleConnect() {
            string oradb = "Data Source=RA00C000;User Id=voshod;Password=voshod;";
            OracleConnection conn = new OracleConnection(oradb);
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "select name from cls";
            cmd.CommandType = CommandType.Text;
            OracleDataReader dr = cmd.ExecuteReader();
            dr.Read();
            System.Diagnostics.Trace.WriteLine(dr.GetString(0));
            //label1.Text = dr.GetString(0);
            conn.Dispose();
        }
    }

    class Directory {
        private string Name;
        private string[][] Content;
    }
}
