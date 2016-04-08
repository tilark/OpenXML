using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace CreateTableInDoc
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateTable(@"D:\Test\wordTest3.docx");
        }
        /// <summary>
        /// 在DOC中创建一个表格，并且能够指定样式。能够指定行的高度，列的宽度，列的宽度能平分整个页面，
        /// </summary>
        /// <param name="fileName"></param>
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
                            new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new BottomBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new LeftBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new RightBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new InsideHorizontalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        },
                        new InsideVerticalBorder()
                        {
                            Val =
                            new EnumValue<BorderValues>(BorderValues.Single),
                            Size = 12
                        }
                        )
                    );

                table.AppendChild<TableProperties>(tblProp);

                TableRow tr = new TableRow();

                TableCell tc1 = new TableCell();

                tc1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = "100%" }));
                tc1.Append(new Paragraph(new Run(new Text("你好！"))));
                tr.Append(tc1);

                //TableCell tc2 = new TableCell(tc1.OuterXml);
                //tr.Append(tc2);
                table.Append(tr);
                //TableRow tbrow2 = new TableRow(tr.OuterXml);
                //table.Append(tbrow2);
                List<string> listTest = new List<string>();
                for(int i = 0; i< 3; i++)
                {
                    listTest.Add("String" + i.ToString());
                }
                CreateOneRowWithText(table, listTest);
                CreateOneRowWithText(table, listTest);

                doc.MainDocumentPart.Document.Body.Append(table);
                
            }
        }
        /// <summary>
        /// 创建一行，列的数量为List中的Count，列中的内容为List中的内容
        /// </summary>
        /// <param name="table"></param>
        /// <param name="listColumnTest"></param>
        public static void CreateOneRowWithText(Table table, List<string> listColumnText)
        {
            TableRow row = new TableRow();
            foreach(var text in listColumnText)
            {
                TableCell cell = new TableCell();
                cell.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Auto  }));
                cell.Append(new Paragraph(new Run(new Text(text))));
                row.Append(cell);
            }
            table.Append(row);

        }
        public static void CreateOpenXMLWordFile(string fileName)
        {
            using (WordprocessingDocument objWordDocument = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart objMainDocumentPart = objWordDocument.AddMainDocumentPart();
                objMainDocumentPart.Document = new Document(new Body());
                Body objBody = objMainDocumentPart.Document.Body;
                //创建一些需要用到的样式,如标题3,标题4,在OpenXml里面,这些样式都要自己来创建的 
                //ReportExport.CreateParagraphStyle(objWordDocument);
                SectionProperties sectionProperties = new SectionProperties();
                PageSize pageSize = new PageSize();
                PageMargin pageMargin = new PageMargin();
                Columns columns = new Columns() { Space = "220" };//720
                DocGrid docGrid = new DocGrid() { LinePitch = 100 };//360
                //创建页面的大小,页距,页面方向一些基本的设置,如A4,B4,Letter, 
                //GetPageSetting(PageSize,PageMargin);

                //在这里填充各个Paragraph,与Table,页面上第一级元素就是段落,表格.
                objBody.Append(new Paragraph());
                objBody.Append(new Table());
                objBody.Append(new Paragraph());

                //我会告诉你这里的顺序很重要吗?下面才是把上面那些设置放到Word里去.(大家可以试试把这下面的代码放上面,会不会出现打开openxml文件有误,因为内容有误)
                sectionProperties.Append(pageSize, pageMargin, columns, docGrid);
                objBody.Append(sectionProperties);

                //如果有页眉,在这里添加页眉.
                //if (IsAddHead)
                //{
                //    //添加页面,如果有图片,这个图片和上面添加在objBody方式有点不一样,这里搞了好久.
                //    //ReportExport.AddHeader(objMainDocumentPart, image);
                //}
                objMainDocumentPart.Document.Save();
            }
        }
    }
}
