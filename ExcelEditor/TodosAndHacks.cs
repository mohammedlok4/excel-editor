using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelEditor
{
    public static class TodosAndHacks
    {
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
                document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (!sheets.Any())
            {
                // The specified worksheet does not exist.
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        public static void AddDateFormat(string excelFilePath)
        {
            string text = "01-31-1947";
            using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(excelFilePath, true))
            {

                // get the stylesheet from the current sheet    
                var stylesheet = spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet;
                // cell formats are stored in the stylesheet's NumberingFormats
                var numberingFormats = stylesheet.NumberingFormats;

                // cell format string               
                const string dateFormatCode = "dd/mm/yyyy";
                // first check if we find an existing NumberingFormat with the desired formatcode
                var dateFormat =
                    numberingFormats.OfType<NumberingFormat>()
                        .FirstOrDefault(format => format.FormatCode == dateFormatCode);
                // if not: create it
                if (dateFormat == null)
                {
                    dateFormat = new NumberingFormat
                    {
                        NumberFormatId = UInt32Value.FromUInt32(164),
                        // Built-in number formats are numbered 0 - 163. Custom formats must start at 164.
                        FormatCode = StringValue.FromString(dateFormatCode)
                    };
                    numberingFormats.AppendChild(dateFormat);
                    // we have to increase the count attribute manually ?!?
                    numberingFormats.Count = Convert.ToUInt32(numberingFormats.Count());
                    // save the new NumberFormat in the stylesheet
                    stylesheet.Save();
                }
                // get the (1-based) index of the dateformat
                var dateStyleIndex = numberingFormats.ToList().IndexOf(dateFormat) + 1;
                var worksheetPart = GetWorksheetPartByName(spreadsheetDoc, "Sheet1");
                Row row1 = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault();
                int counter = 0;
                Cell cell = row1.Elements<Cell>().FirstOrDefault();

                DateTime dateTime = DateTime.Parse(text);
                double oaValue = dateTime.ToOADate();
                cell.CellValue = new CellValue(oaValue.ToString(CultureInfo.InvariantCulture));

                cell.StyleIndex = Convert.ToUInt32(dateStyleIndex);

                worksheetPart.Worksheet.Save();
                spreadsheetDoc.WorkbookPart.WorkbookStylesPart.Stylesheet.Save();
            }
            
        }
    }
}
