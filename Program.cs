using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create("Test.xlsx", SpreadsheetDocumentType.Workbook))
            {
                WriteExcelFile(document);
            }
        }
        private static void WriteExcelFile(SpreadsheetDocument spreadsheet)
        {
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
            spreadsheet.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));
            WorkbookStylesPart workbookStylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>("rIdStyles");
            Stylesheet stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet = stylesheet;
            uint worksheetNumber = 1;
            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            for (int iSheet = 0; iSheet < 5; iSheet++)
            {
                // replace by your Code of datatable creation
                DataTable dt = new DataTable("Table" + iSheet);
                DataRow dr = null;
                dt.Columns.Add("Column1", typeof(string));
                dt.Columns.Add("Column2", typeof(string));
                dt.Columns.Add("Column3", typeof(string));
                dt.Columns.Add("Column4", typeof(string));
                dt.Columns.Add("Column5", typeof(string));
                dt.Columns.Add("Column6", typeof(string));
                dt.Columns.Add("Column7", typeof(string));
                dt.Columns.Add("Column8", typeof(string));
                dt.Columns.Add("Column9", typeof(string));
                for (int iLoop = 0; iLoop < 1000000; iLoop++)
                {
                    dr = dt.NewRow();
                    dr["Column1"] = "Some Random Data for column 1: " + iLoop;
                    dr["Column2"] = "Some Random Data for column 2: " + iLoop;
                    dr["Column3"] = "Some Random Data for column 3: " + iLoop;
                    dr["Column4"] = "Some Random Data for column 4: " + iLoop;
                    dr["Column5"] = "Some Random Data for column 5: " + iLoop;
                    dr["Column6"] = "Some Random Data for column 6: " + iLoop;
                    dr["Column7"] = "Some Random Data for column 7: " + iLoop;
                    dr["Column8"] = "Some Random Data for column 8: " + iLoop;
                    dr["Column9"] = "Some Random Data for column 9: " + iLoop;
                    dt.Rows.Add(dr);
                }
                // end of your Code of datatable creation

                string worksheetName = dt.TableName;
                WorksheetPart newWorksheetPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(newWorksheetPart), SheetId = worksheetNumber, Name = worksheetName };
                sheets.Append(sheet);
                WriteDataTableToExcelWorksheet(dt, newWorksheetPart);
                worksheetNumber++;
            }
            spreadsheet.WorkbookPart.Workbook.Save();
        }
        private static void AppendNumericCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.Number
            });
        }
        private static void AppendTextCell(string cellReference, string cellStringValue, ref OpenXmlWriter writer)
        {
            writer.WriteElement(new Cell
            {
                CellValue = new CellValue(cellStringValue),
                CellReference = cellReference,
                DataType = CellValues.String
            });
        }
        private static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        public static string GetExcelColumnName(int columnIndex)
        {
            char firstChar;
            char secondChar;
            char thirdChar;
            if (columnIndex < 26)
            {
                return ((char)('A' + columnIndex)).ToString();
            }
            if (columnIndex < 702)
            {
                firstChar = (char)('A' + (columnIndex / 26) - 1);
                secondChar = (char)('A' + (columnIndex % 26));

                return string.Format("{0}{1}", firstChar, secondChar);
            }
            int firstInt = columnIndex / 26 / 26;
            int secondInt = (columnIndex - firstInt * 26 * 26) / 26;
            if (secondInt == 0)
            {
                secondInt = 26;
                firstInt = firstInt - 1;
            }
            int thirdInt = (columnIndex - firstInt * 26 * 26 - secondInt * 26);

            firstChar = (char)('A' + firstInt - 1);
            secondChar = (char)('A' + secondInt - 1);
            thirdChar = (char)('A' + thirdInt);
            return string.Format("{0}{1}{2}", firstChar, secondChar, thirdChar);
        }
        private static void WriteDataTableToExcelWorksheet(DataTable dt, WorksheetPart worksheetPart)
        {
            OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart, Encoding.ASCII);
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());

            string cellValue = "";
            int numberOfColumns = dt.Columns.Count;
            bool[] IsNumericColumn = new bool[numberOfColumns];
            bool[] IsDateColumn = new bool[numberOfColumns];

            string[] excelColumnNames = new string[numberOfColumns];
            for (int n = 0; n < numberOfColumns; n++)
                excelColumnNames[n] = GetExcelColumnName(n);
            uint rowIndex = 1;

            writer.WriteStartElement(new Row { RowIndex = rowIndex });
            for (int colInx = 0; colInx < numberOfColumns; colInx++)
            {
                DataColumn col = dt.Columns[colInx];
                AppendTextCell(excelColumnNames[colInx] + "1", col.ColumnName, ref writer);
                IsNumericColumn[colInx] = (col.DataType.FullName == "System.Decimal") || (col.DataType.FullName == "System.Int32") || (col.DataType.FullName == "System.Double") || (col.DataType.FullName == "System.Single");
                IsDateColumn[colInx] = (col.DataType.FullName == "System.DateTime");
            }
            writer.WriteEndElement();

            double cellNumericValue = 0;
            foreach (DataRow dr in dt.Rows)
            {
                ++rowIndex;

                writer.WriteStartElement(new Row { RowIndex = rowIndex });

                for (int colInx = 0; colInx < numberOfColumns; colInx++)
                {
                    cellValue = dr.ItemArray[colInx].ToString();
                    cellValue = ReplaceHexadecimalSymbols(cellValue);

                    if (IsNumericColumn[colInx])
                    {
                        cellNumericValue = 0;
                        if (double.TryParse(cellValue, out cellNumericValue))
                        {
                            cellValue = cellNumericValue.ToString();
                            AppendNumericCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, ref writer);
                        }
                    }
                    else if (IsDateColumn[colInx])
                    {
                        DateTime dtValue;
                        string strValue = "";
                        if (DateTime.TryParse(cellValue, out dtValue))
                            strValue = dtValue.ToShortDateString();
                        AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), strValue, ref writer);
                    }
                    else
                    {
                        AppendTextCell(excelColumnNames[colInx] + rowIndex.ToString(), cellValue, ref writer);
                    }
                }
                writer.WriteEndElement();
            }
            writer.WriteEndElement();
            writer.WriteEndElement();

            writer.Close();
        }

    }
}
