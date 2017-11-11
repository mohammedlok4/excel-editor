using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace ExcelEditor
{
    public static class ExcelEditorServices
    {
        /*
         * rowIndex and columnName act as the starting point for filling in data with Column headings.
         * In an excel prefilled with column names, rowIndex and columnName should still point to the first Column heading. 
         * */
        public static bool InsertPlainData(string connection, string query, string excelFilePath, string sheetName, string columnName, uint rowIndex, bool inlcudeHeadings = true)
        {
            bool success = true;
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(excelFilePath, true))
                {
                    var excelEditor = new ExcelEditor(spreadSheet, sheetName);
                    DataTable dt = excelEditor.GetDataTable(query, connection);
                    excelEditor.ExportDataTable(spreadSheet, columnName, rowIndex, dt, inlcudeHeadings);
                }
                
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
                success = false;
            }
            return success;
        }

        public static bool InsertDataTable(DataTable dt, string excelFilePath, string sheetName, string columnName, uint rowIndex, bool inlcudeHeadings = true)
        {
            bool success = true;
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(excelFilePath, true))
                {
                    var excelEditor = new ExcelEditor(spreadSheet, sheetName);
                    if (excelEditor.SheetPrepared)
                    {
                        excelEditor.ExportDataTable(spreadSheet, columnName, rowIndex, dt, inlcudeHeadings);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                success = false;
            }
            return success;
        }

        public static bool InsertDataTable(DataTable dt, Stream excelFile, string sheetName, string columnName, uint rowIndex, bool inlcudeHeadings = true)
        {
            bool success = true;
            try
            {
                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(excelFile, true))
                {
                    var excelEditor = new ExcelEditor(spreadSheet, sheetName);
                    if (excelEditor.SheetPrepared)
                    {
                        excelEditor.ExportDataTable(spreadSheet, columnName, rowIndex, dt, inlcudeHeadings);
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                success = false;
            }
            return success;
        }

        //No sheet name results in replacing formulas from all sheets
        public static bool ReplaceFormulasWithValues(string excelFile, string sheetName = "")
        {
            bool success = true;
            try
            {
                List<string> sheets;
                if(string.IsNullOrEmpty(sheetName))
                    sheets = GetWorksheetNames(excelFile);
                else
                    sheets = new List<string>() {sheetName};

                using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(excelFile, true))
                {
                    var excelEditor = new ExcelEditor(spreadSheet, sheetName);
                    foreach (string sheet in sheets)
                    {
                        //prepare instance for each sheet
                        if (true)
                        {
                            excelEditor.ReplaceFormulaWithValue(sheet);
                        }
                        
                    }
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
                success = false;
            }
            return success;
        }

        //@todo refactor
        public static List<string> GetTextBoxes(string excelFile, string sheetName)
        {
            List<string> values = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(excelFile, true))
            {
                Sheet sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                if (sheet == null)
                {
                    // The specified worksheet does not exist.
                    return null;
                }

                string relationshipId = sheet.Id.Value;

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

                var ocaElems =
                    worksheetPart.DrawingsPart.WorksheetDrawing.Elements<OneCellAnchor>();
                foreach (OneCellAnchor oneCellAnchor in ocaElems)
                {
                    var shapes = oneCellAnchor.Elements<A.Shape>();
                    foreach (var shape in shapes)
                    {
                        var text = shape.TextBody.InnerText;
                        values.Add(text);
                    }

                }
            }
            return values;
        }

        public static List<string> GetWorksheetNames(string excelFile)
        {
            var sheetNames = new List<string>();
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(excelFile, true))
            {
                IEnumerable<Sheet> sheets = spreadSheet.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                foreach (Sheet sheet in sheets)
                {
                    sheetNames.Add(sheet.Name);
                }
            }
            return sheetNames;
        }
        
    }
}
