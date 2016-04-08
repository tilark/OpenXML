using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
namespace ApplyStyleParagraph
{
    class ApplyStyleToParagraphs
    {
        static void Main(string[] args)
        {
            string fileName = @"D:\Test\wordStyleP.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fileName, true))
            {
                Paragraph p = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>()
                    .ElementAtOrDefault(1);

                if(p == null)
                {
                    throw new ArgumentOutOfRangeException("p", "Paragraph was not found.");
                }

                ApplyStyleToParagraph(doc, "OverAmountTest", "Over Amount Test", p);
            }
        }
        public static  void ApplyStyleToParagraph(WordprocessingDocument doc, string styleid, string stylename, Paragraph p)
        {
            //如果段落中没有ParagraphProperties，创建一个新的
            if (p.Elements<ParagraphProperties>().Count() == 0)
            {
                p.PrependChild<ParagraphProperties>(new ParagraphProperties());

            }
            //获取ParagraphProperties
            ParagraphProperties pPr = p.Elements<ParagraphProperties>().First();
            //获取Styles部件
            StyleDefinitionsPart part = doc.MainDocumentPart.StyleDefinitionsPart;
            //如果Styles部件不存在，创建一个
            if (part == null)
            {
                part = AddStylePartToPackage(doc);
                AddNewStyle(part, styleid, stylename);
            }
            else
            {
                //查看该style是否在文档中，如果没有，添加
                if(IsStyleIdInDocument(doc, styleid) != true)
                {
                    //继续用styleName寻找
                    string styleidFromName = GetStyleIdFromStyleName(doc, stylename);
                    if (styleidFromName == null)
                    {
                        AddNewStyle(part, styleid, stylename);
                    }
                    else
                    {
                        styleid = styleidFromName;
                    }
                }
            }
            pPr.ParagraphStyleId = new ParagraphStyleId() { Val = styleid };
        }

        private static string GetStyleIdFromStyleName(WordprocessingDocument doc, string stylename)
        {
            StyleDefinitionsPart stylePart = doc.MainDocumentPart.StyleDefinitionsPart;
            string styleId = stylePart.Styles.Descendants<StyleName>()
                 .Where(s => s.Val.Value.Equals(stylename) &&
                 (((Style)s.Parent).Type == StyleValues.Paragraph))
                 .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
            return styleId;
                
        }

        private static bool IsStyleIdInDocument(WordprocessingDocument doc, string styleid)
        {
            Styles s = doc.MainDocumentPart.StyleDefinitionsPart.Styles;

            int n = s.Elements<Style>().Count();
            if (n == 0)
                return false;

            //寻找Styleid
            Style style = s.Elements<Style>()
                .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                .FirstOrDefault();
            if (style == null)
                return false;
            return true;
        }

        private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename)
        {
            Styles styles = styleDefinitionsPart.Styles;
            //创建一个新的段落style 并且配置一些新属性
            Style style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true
            };
            StyleName styleName1 = new StyleName() { Val = stylename };
            BasedOn basedOn1 = new BasedOn() { Val = "Normal" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "Normal" };
            style.Append(styleName1);
            style.Append(basedOn1);
            style.Append(nextParagraphStyle1);

            //字体
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color() { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts() { Ascii = "Lucida Console", EastAsia = "宋体" };
            Italic italic1 = new Italic();
            //字体大小
            FontSize fontSize1 = new FontSize() { Val = "24" };
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            style.Append(styleRunProperties1);
            styles.Append(style);
        }

        private static StyleDefinitionsPart AddStylePartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Styles root = new Styles();
            root.Save(part);
            return part;
        }
    }
}
