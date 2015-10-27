using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;

namespace DocxText {
    class Program {

        const int RESULT_COLUMN = 13;
        const string PROGRAM_DIR = @"e:\akriko_converter\!_dev\";

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
                    checkResult = CheckAppeal(appealsTable.Rows[rowIndex]);
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

            document.SaveAs(PROGRAM_DIR + "Ошибки конвертации " + (DateTime.Now).ToString( "yyyy-MM-dd-H-mm-ss" ) + ".docx");
            // System.Diagnostics.Trace.WriteLine((DateTime.Now).ToString( "yyyy-MM-dd-H-mm-ss" ));
        }

        static string CheckAppeal(Row row) {
            string result = "";

            for (int colIndex = 0; colIndex < row.Cells.Count - 1; colIndex++) {
                Cell c = row.Cells[colIndex];
                for (int i = 0; i < c.Paragraphs.Count; i++) {

                    if (c.Paragraphs[i].Text.Trim() != "") {
                        result += c.Paragraphs[i].Text + "; ";
                    }
                }
                result += " | ";
            }
            return result;
        }
    }
}
