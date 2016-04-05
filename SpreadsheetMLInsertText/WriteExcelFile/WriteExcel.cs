using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
namespace WriteExcelFile
{
    class WriteExcel
    {
        public class WriteData
        {
            public WriteData()
            {
                this.BirthDate = DateTime.Now;
            }
            public int Id { get; set; }
            public string Name { get; set; }
            public int Age { get; set; }
            public DateTime BirthDate { get; set; }
        }
        public class WriteDataMul
        {

            public WriteData data1 { get; set; }
            public WriteData data2 { get; set; }
            public WriteData data3 { get; set; }
            public WriteData data4 { get; set; }
            public WriteData data5 { get; set; }
            public WriteData data6 { get; set; }


        }


        static void Main(string[] args)
        {
            WriteData a = new WriteData();
            var b = a?.BirthDate?.Year ;
            ? : 
            List<WriteData> listText = new List<WriteData>();
            for (int i = 0; i < 3; i++)
            {
                WriteData data = new WriteData { Id = i, Name = "Liu", Age = 20 + i };
                listText.Add(data);
            }
            List<String> headColumn = new List<String>();
            headColumn.Add("编号");
            headColumn.Add("姓名");
            headColumn.Add("年龄");
            headColumn.Add("出生日期");
            //WriteExcelFile<WriteData>(@"E:\TestOpenXML\writeTest2.xlsx", listText, headColumn);
            List<int> mulListString = new List<int>();
            for (int i = 0; i < 100; i++)
            {
                mulListString.Add(i);
            }
            WriteExcelFile<int>(@"E:\TestOpenXML\writeTest3.xlsx", mulListString, headColumn);

        }
        /// <summary>
        /// 将一系列数据写入到指定Excel文件中
        /// </summary>
        /// <typeparam name="TData"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="ListText"></param>
        public static void WriteExcelFile<T>(string fileName, List<T> listText, List<String> headColumn)
        {
            //创建Excel文件
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(fileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                WorksheetPart worksheetPart = CreateWorksheetPart(spreadsheetDocument);
                Worksheet worksheet = worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                var shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();

                //先插入列标题

                uint rowIndex = 1;
                InsertHeadColumn(rowIndex, worksheetPart, headColumn, shareStringPart);
                rowIndex = 2;
                //InsertTextToCellWithColumn(rowIndex, worksheetPart, listText, shareStringPart);
                TestInsertTextToCellWithColumn(rowIndex, worksheetPart, listText, shareStringPart);



            }
        }
        private static void InsertTextToCellWithColumn<T>(uint rowIndex, WorksheetPart worksheetPart, List<T> ListText, SharedStringTablePart shareStringPart)
        {
            if (worksheetPart == null || ListText == null || shareStringPart == null)
            {
                return;
            }
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            foreach (var Text in ListText)
            {
                Row row;
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
                int columnIndex = Asc("A");
                string columnName = String.Empty;
                string cellReference = null;
                //将数据写入该列名中
                //判断该数据是不是Text，如果是数值，可直接加入，但在用户信息处，都是Text
                var typeText = Text.GetType();

                foreach (var prop in typeText.GetProperties())
                {
                    //已经是行数据了，可将值插入到Row中
                    cellReference = columnName + Chr(columnIndex) + rowIndex;
                    Cell newCell = new Cell() { CellReference = cellReference };
                    row.AppendChild(newCell);
                    worksheet.Save();
                    //需对传入的Text进行判断，如果是数值型的直接填入，如果是字符串，再填入到SharedStringItem
                    var value = prop.GetValue(Text);
                    if (String.Compare(prop.PropertyType.Name, "String", true) == 0)
                    {
                        var index = InsertSharedStringItem(value.ToString(), shareStringPart);
                        value = index.ToString();
                        newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    }
                    newCell.CellValue = new CellValue(value.ToString());
                    worksheetPart.Worksheet.Save();
                    columnIndex++;
                    if (columnIndex > Asc("Z"))
                    {
                        columnIndex = Asc("A");
                        columnName += "A";
                    }
                }
                rowIndex++;
            }
        }
        private static void TestInsertTextToCellWithColumn<T>(uint rowIndex, WorksheetPart worksheetPart, List<T> ListText, SharedStringTablePart shareStringPart)
        {
            if (worksheetPart == null || ListText == null || shareStringPart == null)
            {
                return;
            }
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row row;
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
            char columnIndex = 'A';
            string columnName = String.Empty;
            string cellReference = null;
            foreach (var Text in ListText)
            {

                //将数据写入该列名中
                //判断该数据是不是Text，如果是数值，可直接加入，但在用户信息处，都是Text
                var typeText = Text.GetType();

                
                //已经是行数据了，可将值插入到Row中
                cellReference = columnName + columnIndex.ToString() + rowIndex;
                Cell newCell = new Cell() { CellReference = cellReference };
                row.AppendChild(newCell);
                worksheet.Save();
                //需对传入的Text进行判断，如果是数值型的直接填入，如果是字符串，再填入到SharedStringItem
                var value = Text.ToString(); ;
                if (String.Compare(Text.GetType().Name, "String", true) == 0)
                {
                    var index = InsertSharedStringItem(value.ToString(), shareStringPart);
                    value = index.ToString();
                    newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }
                newCell.CellValue = new CellValue(value.ToString());
                worksheetPart.Worksheet.Save();
                columnIndex++;
                if (columnIndex > Asc("Z"))
                {
                    //A 之后为AA ,AZ然后是BA,ZZ后面是AAA
                    columnIndex = Asc("A");
                    //先把column加1，由A变为B，再赋值
                    var charByte = new Char[5] ;
                    
                    for (int i = 0; i < columnName.Length; i++)
                    {
                        var columnChar = columnName[i];
                        
                        if (Asc(columnChar.ToString()) > Asc("Z"))
                        {
                            //columnName[0]需变为A
                            charByte[i] = Char.Parse("A");
                        }
                    }
                    columnName += "A";
                }


            }
            rowIndex++;
        }
        private static void InsertHeadColumn(uint rowIndex, WorksheetPart worksheetPart, List<String> headColumn, SharedStringTablePart shareStringPart)
        {
            if (worksheetPart == null || headColumn == null)
            {
                return;
            }
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            Row row;
            row = new Row() { RowIndex = rowIndex };
            sheetData.Append(row);
            int columnIndex = Asc("A");
            string columnName = String.Empty;
            string cellReference = null;
            foreach (var text in headColumn)
            {
                //将数据写入该列名中
                //已经是行数据了，可将值插入到Row中
                cellReference = columnName + Chr(columnIndex) + rowIndex;
                Cell newCell = new Cell() { CellReference = cellReference };
                row.AppendChild(newCell);
                worksheet.Save();
                //填入到SharedStringItem
                var index = InsertSharedStringItem(text, shareStringPart);
                newCell.CellValue = new CellValue(index.ToString());
                newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                worksheetPart.Worksheet.Save();
                columnIndex++;
                if (columnIndex > Asc("Z"))
                {
                    columnIndex = Asc("A");
                    columnName += "A";
                }
            }
        }
        private static WorksheetPart CreateWorksheetPart(SpreadsheetDocument spreadsheetDocument)
        {
            //WorksheetPart worksheetPart = null;
            #region ini spreadsheetDocument
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();
            #endregion
            return worksheetPart;
        }
        //字符转ASCII码：
        public static int Asc(string character)
        {
            if (character.Length == 1)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                int intAsciiCode = (int)asciiEncoding.GetBytes(character)[0];
                return (intAsciiCode);
            }
            else
            {
                throw new Exception("Character is not valid.");
            }

        }

        //ASCII码转字符：

        public static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
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
        public static void CreateSpreadsheetWorkbook(SpreadsheetDocument spreadsheetDocument)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

        }

    }
}
