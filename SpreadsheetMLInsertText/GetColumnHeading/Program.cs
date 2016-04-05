using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace GetColumnHeading
{
    class Program
    {
        static void Main(string[] args)
        {
        }
        public static string GetColumnHeading(string docName, string worksheetName, string cellName)
        {
            using (SpreadsheetDocument spreadSheetDoc = SpreadsheetDocument.Open(docName, false))
            {
                var sheets = spreadSheetDoc.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == worksheetName);
                if(sheets.Count() == 0)
                {
                    return null;
                }
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDoc.WorkbookPart.GetPartById(sheets.First().Id);
                string columnName = GetColumnName(cellName);
                var cells = worksheetPart.Worksheet.Descendants<Cell>().Where(c => string.Compare(GetColumnName(c.CellReference.Value), columnName, true) == 0)
                    .OrderBy(r => GetRowIndex(r.CellReference));
                if(cells.Count() == 0)
                {
                    return null;
                }
                Cell headCell = cells.First();
                if(headCell.DataType != null && headCell.DataType.Value == CellValues.SharedString)
                {
                    SharedStringTablePart shareStringPart = spreadSheetDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    var items = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ToArray();
                    return items[int.Parse(headCell.CellValue.Text)].InnerText;
                }
                else
                {
                    return headCell.CellValue.Text;
                }
            }
        }

        private static object GetRowIndex(StringValue cellReference)
        {
            throw new NotImplementedException();
        }

        private static string GetColumnName(string cellName)
        {
            throw new NotImplementedException();
        }
    }
}
