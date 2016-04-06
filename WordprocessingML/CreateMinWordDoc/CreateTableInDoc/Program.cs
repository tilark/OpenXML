using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace CreateTableInDoc
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateTable(@"D:\Test\test2.docx");
        }
        public static void CreateTable(string fileName)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, true))
            {
                Table table = new Table();
                TableProperties tblProp = new TableProperties(
                    new TableBorders(
                        new TopBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new BottomBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new LeftBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new RightBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new InsideHorizontalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        },
                        new InsideVerticalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Dashed),
                            Size = 24
                        }
                        )
                    );

                table.AppendChild<TableProperties>(tblProp);

                TableRow tr = new TableRow();

                TableCell tc1 = new TableCell();

                tc1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
                tc1.Append(new Paragraph(new Run(new Text("你好！"))));
                tr.Append(tc1);

                TableCell tc2 = new TableCell(tc1.OuterXml);
                tr.Append(tc2);
                table.Append(tr);
                TableRow tbrow2 = new TableRow(tr.OuterXml);
                table.Append(tbrow2);
                doc.MainDocumentPart.Document.Body.Append(table);
                
            }
        }
    }
}
