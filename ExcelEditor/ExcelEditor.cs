using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ExcelEditor
{
    public class ExcelEditor
    {
        public enum DataTypes
        {
            String,
            DateTime,
            Number,
        };

        private SpreadsheetDocument _spreadSheet;
        private string _sheetName;
        private WorksheetPart _worksheetPart;
        private List<string> _columnHeaders = new List<string>();
        //@TODO refactor
        public int _dateStyleIndex = -1;

        public bool SheetPrepared = false;
        public bool Log { get; set; }

        public ExcelEditor(SpreadsheetDocument sd, string sheetName)
        {
            try
            {
                _spreadSheet = sd;
                _sheetName = sheetName;
                // Open the document for editing.
                _worksheetPart = GetWorksheetPartByName(_spreadSheet, _sheetName);
                _columnHeaders.Clear();

                CellFormat cellFormat = new CellFormat()
                {
                    NumberFormatId = (UInt32Value)14U,
                    FontId = (UInt32Value)0U,
                    FillId = (UInt32Value)0U,
                    BorderId = (UInt32Value)0U,
                    FormatId = (UInt32Value)0U,
                    ApplyNumberFormat = true
                };

                sd.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);
                _dateStyleIndex = sd.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Count() - 1;

                if (_worksheetPart != null)
                {
                    SheetPrepared = true;
                }
                Log = true;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
            }
        }

        public void ExportDataTable(SpreadsheetDocument spreadSheet, string columnName, uint rowIndex, DataTable dt,
            bool includeHeadings = true)
        {
            //Columns headers
            int count = dt.Columns.Count;
            string[] columnNames = new string[count];
            string[] dataTypes = new string[count];

            string col = columnName;
            for (int i = 0; i < count; i++)
            {
                dataTypes[i] = dt.Columns[i].DataType.Name;
                columnNames[i] = dt.Columns[i].ColumnName;
                _columnHeaders.Add(dt.Columns[i].ColumnName);
                //Header will only contain strings
                if (includeHeadings)
                    UpdateCell(columnNames[i], rowIndex, col, ExcelEditor.DataTypes.String);

                col = Helper.NextLetter(col);

            }

            //Dump the data
            for (int i = 0; i < count; i++)
            {
                ExcelEditor.DataTypes dtTmp;

                if (dataTypes[i] == "String")
                {
                    dtTmp = ExcelEditor.DataTypes.String;
                }
                else if (dataTypes[i] == "DateTime")
                {
                    dtTmp = ExcelEditor.DataTypes.DateTime;
                }
                else
                {
                    dtTmp = ExcelEditor.DataTypes.Number;
                }
                List<string> columnData = dt.AsEnumerable().Select(x => x[i].ToString()).ToList();

                uint rowCounter = rowIndex + 1;
                foreach (string column in columnData)
                {
                    UpdateCell(column, rowCounter, columnName, dtTmp);
                    rowCounter++;
                }
                columnName = Helper.NextLetter(columnName);
            }
            SaveChanges();
        }


        public void UpdateCell(string text, uint rowIndex, string columnName, DataTypes type)
        {

            if (!SheetPrepared)
            {
                return;
            }

            Cell cell = GetCell(_worksheetPart.Worksheet, columnName, rowIndex);
            UpdateCell(cell, type, text);
        }


        private void UpdateCell(Cell cell, DataTypes type, string text)
        {
            if (type == DataTypes.String)
            {
                //cell.CellValue = new CellValue(text);
                cell.DataType = CellValues.SharedString;

                if (!_spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
                {
                    _spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                var sharedStringTablePart = _spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                if (sharedStringTablePart.SharedStringTable == null)
                {
                    sharedStringTablePart.SharedStringTable = new SharedStringTable();
                }
                //Iterate through shared string table to check if the value is already present.
                foreach (SharedStringItem ssItem in sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
                {
                    if (ssItem.InnerText == text)
                    {
                        cell.CellValue = new CellValue(ssItem.ElementsBefore().Count().ToString());
                        SaveChanges();
                        return;
                    }

                }
                // The text does not exist in the part. Create the SharedStringItem.
                var item = sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                cell.CellValue = new CellValue(item.ElementsBefore().Count().ToString());
            }
            else if (type == DataTypes.Number)
            {
                cell.CellValue = new CellValue(text);
                cell.DataType = CellValues.Number;
            }
            else if (type == DataTypes.DateTime)
            {

                cell.DataType = CellValues.Number;
                cell.StyleIndex = Convert.ToUInt32(_dateStyleIndex);

                DateTime dateTime = DateTime.Parse(text);
                double oaValue = dateTime.ToOADate();
                cell.CellValue = new CellValue(oaValue.ToString(CultureInfo.InvariantCulture));
            }

            // Save the worksheet.
            SaveChanges();
        }

        internal void ReplaceFormulaWithValue(string sheetName)
        {
            var formulaDict = new Dictionary<string, string>();
            StringBuilder logStringBuilder = new StringBuilder();

            var worksheetPart = GetWorksheetPartByName(_spreadSheet, sheetName);
            _worksheetPart = worksheetPart;

            CalculationChainPart calculationChainPart = _spreadSheet.WorkbookPart.CalculationChainPart;
            if (calculationChainPart == null)
                return;

            CalculationChain calculationChain = calculationChainPart.CalculationChain;
            var calculationCells = calculationChain.Elements<CalculationCell>().ToList();

            foreach (Row row in worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellValue != null)
                    {
                        Console.WriteLine(cell.CellReference);
                    }
                    if (cell.CellFormula != null &&
                        cell.CellValue != null)
                    {

                        string cellRef = cell.CellReference;
                        string formula = cell.CellFormula.InnerText;

                        if (!string.IsNullOrEmpty(cell.CellFormula.SharedIndex))
                        {
                            if (!formulaDict.ContainsKey(cell.CellFormula.SharedIndex.InnerText))
                            {
                                formulaDict.Add(cell.CellFormula.SharedIndex.InnerText, cell.CellFormula.InnerText);
                            }
                        }
                        //
                        if (formula == "" && cell.CellFormula.SharedIndex != null)
                        {
                            string tmp;
                            if (formulaDict.TryGetValue(cell.CellFormula.SharedIndex.InnerText, out tmp))
                            {
                                formula = "Shared " + tmp;
                            }
                        }

                        CalculationCell calculationCell =
                            calculationCells.Where(c => c.CellReference == cellRef).FirstOrDefault();
                        //CalculationCell calculationCell = calculationChain.Elements<CalculationCell>().Where(c => c.CellReference == cell.CellReference).FirstOrDefault();

                        string value = cell.CellValue.InnerText;
                        UpdateCell(cell, DataTypes.String, value);

                        cell.CellFormula.Remove();
                        calculationCell.Remove();
                        //Try
                        calculationCells.Remove(calculationCell);
                        //Log
                        if (Log)
                        {
                            string log = string.Format("Cell: {0} | Sheet: {1} | Formula: {2} | Value: {3}", cellRef,
                                sheetName, formula, value);
                            logStringBuilder.Append(log);
                            logStringBuilder.Append(Environment.NewLine);
                        }

                    }
                    if (calculationCells.Count == 0)
                    {
                        //delete calcCalutions.xml
                        _spreadSheet.WorkbookPart.DeletePart(calculationChainPart);
                    }
                }

            }
            //Log
            if (Log)
            {
                File.AppendAllText("Log.log", logStringBuilder.ToString());
            }
            //SaveChanges();
        }

        public void SaveChanges()
        {
            _worksheetPart.Worksheet.Save();
        }

        private WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
                document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (!sheets.Any())
            {
                // The specified worksheet does not exist.
                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart) document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);
            if (row == null)
                return null;

            string cellReference = columnName + rowIndex;

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex))
            {
                return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() {CellReference = cellReference};

                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }


        // Given a worksheet and a row index, return the row.
        private Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            Row row = null;

            var res = worksheet.GetFirstChild<SheetData>().Elements<Row>();
            if (worksheet.GetFirstChild<SheetData>().Elements<Row>().Any(r => r.RowIndex == rowIndex))
            {
                row = worksheet.GetFirstChild<SheetData>().Elements<Row>().First(r => r.RowIndex == rowIndex);
            }

            if (row == null)
            {
                //Add the new row to sheet data
                var sheetData = worksheet.GetFirstChild<SheetData>();
                uint lastRowIdx = 0;

                var lastOrDefault = sheetData.Elements<Row>().LastOrDefault();
                if (lastOrDefault != null)
                {
                    lastRowIdx = lastOrDefault.RowIndex;
                }

                row = new Row()
                {
                    RowIndex = rowIndex
                };
                var lastRow = sheetData.Elements<Row>().LastOrDefault();
                sheetData.InsertAfter<Row>(row, lastRow);
            }

            return row;
        }

        public DataTable GetDataTable(string query, string myConnection)
        {
            var dt = new DataTable();

            using (SqlConnection connection = new SqlConnection(myConnection))
            {
                using (var adapter = new SqlDataAdapter(query, connection))
                {
                    try
                    {
                        adapter.SelectCommand = new SqlCommand(query, connection);
                        adapter.Fill(dt);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }
            }
            return dt;
        }

        //Test
        public void DateTimeTest()
        {
            uint rowIndex = 2;
            string date = "2016-01-10";
            //string date = "DATE";

            //Read in NumFmt
            var numFmts = _spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats;

            foreach(var nf in numFmts)
            {
                string ox = nf.OuterXml;
                var xmlElem = XElement.Parse(ox);
                string num = xmlElem.Attribute("numFmtId").Value;
                string formatCode = xmlElem.Attribute("formatCode").Value;
                Utility.ParseAndAddNumFmt(formatCode);
            }

            for (int i = 0; i < 10; i++)
            {
                Cell cell = GetCell(_worksheetPart.Worksheet, "F", rowIndex);
                var d = cell.CellValue.Text;

                int si;
                if (int.TryParse(cell.StyleIndex.ToString(), out si))
                {
                    CellFormat cellFormat = (CellFormat)_spreadSheet.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(si);
                    var numFmt = cellFormat.NumberFormatId;
                    if (numFmt == 14)
                    {
                        double dateVal = Double.Parse(d);
                        var dateTmp = DateTime.FromOADate(dateVal);
                        string dateString = dateTmp.ToString("M-d-yy");
                        Console.WriteLine(dateString);
                    }
                    else if(numFmt == 166)
                    {
                        double dateVal = Double.Parse(d);
                        var dateTmp = DateTime.FromOADate(dateVal);
                        string dateString = dateTmp.ToString("dddd, MMMM dd, yyyy");
                        Console.WriteLine(dateString);
                    }

                } 
                rowIndex++;

            }
            SaveChanges();
        }

       
    }
}
