
using DocumentFormat.OpenXml.Packaging;
using System;

namespace ExcelEditor
{
    class Program
    {
        static void Main(string[] args)
        {
            string connection = @"Data Source=LAPTOP-N1RJNKED\SQLEXPRESS;Initial Catalog=RedactionDummy;Integrated Security=True";
            //string query = "SELECT TOP 100 [EmpData].first_name, [EmpData].last_name, [EmpData].company_name, [EmpData].city, [EmpData].phone1, [EmpData].email FROM EmpData";
            string query = "SELECT * FROM Table_1";
            string excelFilePath = "Chart.xlsx";

            string sheetName = "Tabelle1";
            string columnIndex = "B";
            uint rowIndex = 2;
            //var result = ExcelEditorServices.InsertPlainData(connection, query, excelFilePath, sheetName, columnIndex, rowIndex, false);

            //Console.WriteLine("Result: {0}", result.ToString());

            //var result2 = ExcelEditorServices.GetTextBoxes(excelFilePath, sheetName);
            //var result2 = ExcelEditorServices.GetTextBoxes(excelFilePath, sheetName);

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(excelFilePath, true))
            {
                var ee = new ExcelEditor(document, sheetName);
                ee.DateTimeTest();
            }


            Console.WriteLine("Program executed...");
            Console.ReadKey();
        }

        
    }
}
