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

        public PdfDocument ToPdfOLD(XDocument xdoc)
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
            string fontName = "T1";

            var paras = xdoc.Descendants(w + "p").ToList();


            CreateTestPdf(@"c:\dumpzone\test.pdf");

//            StreamWriter output = new StreamWriter(@"c:\dumpzone\output.pdf");
            StringWriter output = new StringWriter();
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
                        //var run = new TextObject(textNode.TabPos, yPos, textNode.Text,Courier, fontSize, textNode.Justification);
                        //output.WriteLine(run);
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
            return null;    //output.ToString();
        }

        private void CreateTestPdf(string filename)
        {
            var pdf = new PdfDocument();
            PageObject page = new PageObject();
            pdf.AddPage(page, new PageDescription(612, 792, 10, 10, 10, 10));

            FileStream output = new FileStream(filename, FileMode.Create);
            FontObject Courier = new FontObject("CourierNew");
            page.AddFont(Courier);

            ContentObject contentObj = new ContentObject();
            page.AddContent(contentObj);
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
