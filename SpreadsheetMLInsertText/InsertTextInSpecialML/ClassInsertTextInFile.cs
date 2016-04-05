using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace InsertTextInSpecialML
{
    class ClassInsertTextInFile
    {
        static void Main(string[] args)
        {
            InsertTextInFile(@"E:\TestOpenXML\test4.xlsx", "Sheet1", "A1" ,"Inserted Text");
            InsertTextInFile(@"E:\TestOpenXML\test4.xlsx", "Sheet1", "B2", "Inserted Text");
            InsertTextInFile(@"E:\TestOpenXML\test4.xlsx", "Sheet1", "C1", "Inset");
            InsertTextInFile(@"E:\TestOpenXML\test4.xlsx", "Sheet2", "C1", "Inset");


        }
        /// <summary>
        /// 将文本插入到指定Excel文件的指定工作簿中的指定单元格中.
        /// </summary>
        /// <param name="docName">指定Excel文件名.</param>
        /// <param name="sheetName">指定工作簿名.</param>
        /// <param name="cellName">指定单元格</param>
        /// <param name="text">需插入的文本</param>
        public static void InsertTextInFile(string docName, string sheetName, string cellName, string text)
        {
            //判断输入的参数是否合法
            InsertText(docName, sheetName, cellName, text);
        }
        /// <summary>
        /// Inserts the text.
        /// </summary>
        /// <param name="docName">Name of the document.</param>
        /// <param name="text">The text.</param>
        public static void InsertText(string docName, string sheetName, string cellName, string text)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                var sheets = spreadSheet.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);
                WorksheetPart worksheetPart = null;
                if (sheets.Count() == 0)
                {
                    //未找到制定的工作簿
                    //找到指定的工作簿，如果没有找到，可以选择新建一个工作簿
                    //Insert a new worksheet
                    worksheetPart = InsertWorksheet(spreadSheet.WorkbookPart);
                    //return;
                }
                else
                {
                    worksheetPart = (WorksheetPart)spreadSheet.WorkbookPart.GetPartById(sheets.First().Id);
                }
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                int index = InsertSharedStringItem(text, shareStringPart);
                //将指定的单元格插入到工作簿中
                var rowIndex = GetRowIndex(cellName);
                var columnName = GetColumnName(cellName);
                Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);

                cell.CellValue = new CellValue(index.ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                worksheetPart.Worksheet.Save();
            }
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            //if there is not a cell with the specified column name, insert one.
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                //Cells must be in sequential order according to CellReference.
                //Determine where to insert the new cell
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (String.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }
                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }

        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            //Get a unique ID for the new sheet
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        // Given a cell name, parses the specified cell to get the row index.
        private static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
        // Given a cell name, parses the specified cell to get the column name.
        private static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
    }
}
