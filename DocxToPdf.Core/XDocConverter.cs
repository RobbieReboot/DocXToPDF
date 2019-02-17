using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace DocxToPdf.Core
{
    public class XDocConverter
    {
        public string ToPdf(XDocument xdoc)
        {
            var reader = xdoc.Root.CreateReader();
            var nsm = new XmlNamespaceManager(reader.NameTable);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace pt14 = "http://powertools.codeplex.com/2011";
            //nsm.AddNamespace(w.NamespaceName, w.ToString());            // "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            //nsm.AddNamespace(pt14.NamespaceName, pt14.ToString());      //"pt14", "http://powertools.codeplex.com/2011");
            nsm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            nsm.AddNamespace("pt14", "http://powertools.codeplex.com/2011");

            int paraNum = 0;
            int yPos = 0;
            int fontSize = 9;
            double hScale = 1.1;
            double vScale = 1.6;
            var paras = xdoc.Descendants(w + "p").ToList();

            StringWriter output = new StringWriter();

            InitPdf();
            

//            StreamWriter output = new StreamWriter(@"c:\dumpzone\output.pdf");
            output.WriteLine(GetHeader());
            output.WriteLine(
        @"4 0 obj
<< /Length 62 >>
stream");
            foreach (var para in paras)
            {
                var tabs = para.XPathSelectElements("w:pPr/w:tabs/w:tab", nsm).ToList();
                var rNodes = para.XPathSelectElements("w:r", nsm).Skip(1);
                var tNodes = para.XPathSelectElements("w:r/w:t", nsm);

                var anytext = tNodes.Any(tn => String.IsNullOrWhiteSpace(tn.Value));
                var alltext = tNodes.All(tn => String.IsNullOrWhiteSpace(tn.Value));
                //	$"TabCount  = {tabs.Count}\n<r> count = {rNodes.Count()}\n<t> count = {tNodes.Count()}\n".Dump($"Paragraph {paraNum}");

                var matchedTabNodes = new List<Tuple<string, double, string>>()
                          .Select(t => new { Text = string.Empty, TabPos = double.MinValue, Justification = String.Empty }).ToList();
                var tabNo = 0;
                var tabCount = tabs.Count();
                foreach (var rnode in rNodes)
                {
                    var newNode = new
                    {
                        Text = rnode.XPathSelectElement("w:t", nsm)?.Value,
                        TabPos = (double.Parse(tabs[tabNo].Attribute(w + "pos").Value) / 20) * hScale,  //(fontSize*1.6),
                        Justification = tabs[tabNo].Attribute(w + "val").Value
                    };
                    matchedTabNodes.Add(newNode);
                    if (string.IsNullOrEmpty(newNode.Text))
                        tabNo++;
                    if (tabNo >= tabs.Count) break;
                }
                var final = matchedTabNodes.Where(n => !String.IsNullOrEmpty(n.Text)).ToList();
                foreach (var textNode in final)
                {
                    //Null entry means skip a line.
                    if (textNode == null)
                    {
                        yPos += (int)(fontSize * vScale);
                        continue;
                    }
                    //no text? skip object output!
                    if (!string.IsNullOrEmpty(textNode.Text))
                    {
                        var run = TextObject(textNode.TabPos, yPos, textNode.Text, fontSize, textNode.Justification);
                        output.WriteLine(run);
                    }
                }


                //		tabNo = 0;
                //				var result = rNodes.Select((rnode, idx) => new {
                //						Text = rnode.XPathSelectElement("w:t", nsm)?.Value,
                //						TabPos = double.Parse(tabs[tabNo].Attribute(w + "pos").Value) / 20,
                //						Justification = tabs[tabNo += (string.IsNullOrEmpty(rnode.XPathSelectElement("w:t", nsm)?.Value) ? 1 : 0)].Attribute(w + "val").Value,
                //					});
                //		result.Dump("Linqed");
                //Get previous TAB value
                //var previousTab = textNodes[1].XPathEvaluate("string(preceding-sibling::*[1]/w:tab/@pt14:TabWidth)",nsm).Dump("Previous Run TabWidth") ;
                //var tabPos = double.Parse(textNodes[1].PreviousNode.XPathSelectElement("w:tab",nsm).Attribute(pt14 + "TabWidth")?.Value) / 20;

                paraNum++;
                yPos += (int)(fontSize * vScale);
            }
            output.WriteLine(
        @"endstream
endobj
");
            output.WriteLine(GetTrailer());
            output.Flush();
            output.Close();
            return output.ToString();
        }

        private void InitPdf()
        {
            FileStream output = new FileStream(@"c:\Dumpzone\pdfgen.pdf", FileMode.Create);
            CatalogDict catalogDict = new CatalogDict();
            PageTreeDict pageTreeDict = new PageTreeDict();
            FontDict Courier = new FontDict();
            InfoDict infoDict = new InfoDict();
            Courier.CreateFontDict("T1", "Courier");
            infoDict.SetInfo("pdftest", "Rob", "3Squared");

            Utility pdfUtility = new Utility();

            int size = 0;
            output.Write(pdfUtility.GetHeader("1.5", out size));
            output.Flush();
            output.Close();

            PageDict page = new PageDict();
            ContentDict content = new ContentDict();
            PageSize pSize = new PageSize(612, 792);
            pSize.SetMargins(10, 10, 10, 10);
            page.CreatePage(pageTreeDict.objectNum, pSize);
            pageTreeDict.AddPage(page.objectNum);
            page.AddResource(Courier, content.objectNum);
          
            //AddRow(false, 10, "T1", align, "First Column", "Second Column");
            //textAndtable.AddRow(false, 10, "T1", align, "Second Row", "Second Row");
            content.SetStream(TextObject(0,0,"BOLLOX",12,"left"));

            var file = new FileStream(@"c:\Dumpzone\pdfgen.pdf", FileMode.Append);
            file.Write(page.GetPageDict(file.Length, out size), 0, size);
            file.Write(content.GetContentDict(file.Length, out size), 0, size);
            file.Write(catalogDict.GetCatalogDict(pageTreeDict.objectNum,
                file.Length, out size), 0, size);
            file.Write(pageTreeDict.GetPageTree(file.Length, out size), 0, size);
            file.Write(Courier.GetFontDict(file.Length, out size), 0, size);
            file.Write(infoDict.GetInfoDict(file.Length, out size), 0, size);
            file.Write(pdfUtility.CreateXrefTable(file.Length, out size), 0, size);
            file.Write(pdfUtility.GetTrailer(catalogDict.objectNum,
                infoDict.objectNum, out size), 0, size);

            file.Flush();
            file.Close();

        }

        public string TextObject(double xPos, int yPos, string txt, int fontSize, string alignment)
        {
            double startX = 0;
            switch (alignment)
            {
                case "left":
                    startX = xPos;
                    break;
                case "center":
                    startX = xPos - (StrLen(txt, fontSize)) / 2;
                    break;
                case "right":
                    startX = xPos - StrLen(txt, fontSize) + 2;
                    break;
            };
            return string.Format("\rBT/{0} {1} Tf\r{2} {3} Td \r({4}) Tj\rET\r",
                "T0" /*fontName*/, fontSize, startX, (720 - yPos), txt);
        }

        private int StrLen(string text, int fontSize)
        {
            char[] cArray = text.ToCharArray();
            int cWidth = 0;
            foreach (char c in cArray)
            {
                cWidth += 500;  //(int)(fontSize*1.6)*20;	//Monospaced font width?
            }
            //div by1000??? 100 seems to work better :/
            //$"{text} - {(cWidth / 100)}".Dump("StrLen Em's");
            return (cWidth / 100);
        }

        public string GetHeader()
        {
            return
        @"%PDF-1.1
%¥±ë

1 0 obj
  << /Type /Catalog
     /Pages 2 0 R
  >>
endobj

2 0 obj
  << /Type /Pages
     /Kids [3 0 R]
     /Count 1
  >>
endobj

3 0 obj
  <<  /Type /Page
      /Parent 2 0 R
      /Resources
       << /Font
           << /F1
               << /Type /Font
                  /Subtype /Type1
                  /BaseFont /Courier
               >>
           >>
       >>
      /Contents 4 0 R
  >>
endobj";
        }

        public string GetTrailer()
        {
            return
        @"xref
0 5
0000000000 65535 f

trailer
  <<  /Root 1 0 R
      /Size 5
  >>
startxref
609
%%EOF
";
        }

    }





}
