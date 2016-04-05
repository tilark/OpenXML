using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace GetCellValue
{
    class ClassGetCellValue
    {
        static void Main(string[] args)
        {
            string value = GetCellValue(@"E:\TestOpenXML\test4.xlsx", "Sheet1", "D3");
            Console.WriteLine(value);
        }
        /// <summary>
        /// 从指定的文件名和工作簿的单元格中取出数据.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="addressName">Name of the cell address.</param>
        /// <returns>单元格内的内容.</returns>
        public static string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;
            using (SpreadsheetDocument spreadSheetDoc = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadSheetDoc.WorkbookPart;
                Sheet theSheet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(theSheet.Id);
                Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
                if (theCell != null)
                {
                    //输出的格式为XML，如
                    //< x:c r = "A1" >
                    //< x:v > 12.345000000000001 </ x:v >
                    // </ x:c >
                    //Console.WriteLine(theCell.OuterXml);
                    value = theCell.InnerText;
                    //用Excel打开文本，随便写入一个文本0，会显示该单元格无DataType，若无DataType，将其当作普通文本对待
                    if (theCell.DataType != null)
                    {
                        //根据CellType进行文字处理，如果是数字或日期，可直接返回；
                        //如果是文本，需从SharedStringTable中查找；
                        //如果是布尔值，需转换
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                var sharedStringTable = workbookPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();
                                //如果sharedStringTable为null，说明文本已损坏，返回值为索引值
                                if (sharedStringTable != null)
                                {
                                    value = sharedStringTable.SharedStringTable
                                        .ElementAt(int.Parse(value))
                                        .InnerText;
                                }
                                break;
                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }

                }
            }
            return value;
        }
    }
}
