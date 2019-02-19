//
// Rob Hill rob.hill@3Squared.com
//
// Pdf stuff gleaned from:
// http://preserve.mactech.com/articles/mactech/Vol.15/15.09/PDFIntro/index.html
//
// 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace DocxToPdf.Core
{
    /// <summary>
    /// Information source for this from https://www.oreilly.com/library/view/developing-with-pdf/9781449327903/ch01.html
    /// 
    /// </summary>
    public class PdfDocument
    {
        public XRefTableObject xrefTable;
        public uint NextObjectNum { get; set; }
        public int fontIndex = 1;   //numerical Object tag..

        private CatalogObject catalogObj;
        private PageTreeObject pageTreeObj;
        private InfoObject infoObject;
        private List<PageObject> pageObjects;

        public PdfDocument()
        {
            xrefTable = new XRefTableObject();
            pageTreeObj = new PageTreeObject(this);
            catalogObj = new CatalogObject(pageTreeObj,this);
            pageObjects = new List<PageObject>();
        }

        public byte[] CreateXrefTable(long fileOffset, out int size) => xrefTable.RenderBytes(fileOffset, out size);
        public void SetMetadata(InfoObject metadata) => infoObject = metadata;

        public PageObject AddPage(PageExtents pDesc)
        {
            var page = new PageObject(this,pDesc, pageTreeObj.PdfObjectId);
            pageTreeObj.AddPage(page);
            pageObjects.Add(page);
            return page;
        }

        public FontObject AddFont(string fontName)
        {
            return new FontObject(fontName,this);
        }

        public InfoObject AddInfoObject(string title, string author, string company)
        {
            return infoObject = new InfoObject(title,author,company,this);
        }

        public byte[] RenderHeader(string version, out int size)
        {
            //            var headerCode = 0xe2e3cfd3;
            //            string header = $"%PDF-{version}\r%{headerCode}\r\n";
            string header = string.Format("%PDF-{0}\r%{1}\r\n", version, "\x82\x82");
            return GetUTF8Bytes(header, out size);
        }
        /// <summary>
        /// Creates the trailer and return the bytes array
        /// </summary>
        /// <param name="size"></param>
        /// <returns></returns>
        public byte[] RenderTrailer(uint refRoot, uint refInfo, out int size)
        {
            string trailer = null;
            string infoDict = String.Empty;

            if (refInfo > 0)
            {
                infoDict = string.Format("/Info {0} 0 R", refInfo);
            }

            //The sorted array will be already sorted to contain the file offset at the zeroth position
            ObjectXRef objList = xrefTable.ObjectByteOffsets[0];
            trailer = string.Format("trailer\n<</Size {0}/Root {1} 0 R {2}/ID[<5181383ede94727bcb32ac27ded71c68>" +
                                    "<5181383ede94727bcb32ac27ded71c68>]>>\r\nstartxref\r\n{3}\r\n%%EOF\r\n"
                , xrefTable.XRefCount, refRoot, infoDict, xrefTable.FirstByteOffset);

            //xrefTable.ObjectByteOffsets = null;
            //PdfObject.NextObjectNum = 0;

            return GetUTF8Bytes(trailer, out size);
        }

        public static byte[] GetUTF8Bytes(string str, out int size)
        {
            try
            {
                byte[] ubuf = Encoding.Unicode.GetBytes(str);
                Encoding enc = Encoding.GetEncoding("utf-8");
                byte[] abuf = Encoding.Convert(Encoding.Unicode, enc, ubuf);
                size = abuf.Length;
                return abuf;
            }
            catch (Exception e)
            {
                Exception error = new Exception(e.Message + " In Utility.GetUTF8Bytes()");
                throw error;
            }
        }

        public static PdfDocument FromDocX(string xmlFile)
        {
            var xdoc = XDocument.Load(xmlFile);
            return PdfDocument.FromDocX(xdoc);
        }

        public static PdfDocument FromDocX(XDocument xdoc)
        {
            var reader = xdoc.Root.CreateReader();
            var nsm = new XmlNamespaceManager(reader.NameTable);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace pt14 = "http://powertools.codeplex.com/2011";
            nsm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            nsm.AddNamespace("pt14", "http://powertools.codeplex.com/2011");

            int yPos = 0;
            int fontSize = 9;
            double hScale = 1.1;
            double vScale = 1.6;

            var paras = xdoc.Descendants(w + "p").ToList();
            var pdf = new PdfDocument();

            // TODO: Should come from the classes in the doc but for now, its all MonoSpaced.
            FontObject CourierNew = pdf.AddFont("CourierNew");

            // TODO: Should come from the Docs metadata..
            pdf.AddInfoObject("XDoc2Pdf", "RobHill","3Squared");

            // TODO: Point sizes from adobe for A4 - Fixed for now.
            var page = pdf.AddPage(new PageExtents(612, 792, 32, 10, 32, 10));
            // TODO: The text object should add the font to the page IF it's not already added (Possibly). Slower than this but more convenient.

            page.AddFont(CourierNew);

            // TODO: SHould refactor this so the PAGE object is responsible for creation of content object so the parentRef can be set on creation.
            var contentObj = page.AddContentObject();

            foreach (var para in paras)
            {
                var tabs = para.XPathSelectElements("w:pPr/w:tabs/w:tab", nsm).ToList();
                var rNodes = para.XPathSelectElements("w:r", nsm).Skip(1);
                var tNodes = para.XPathSelectElements("w:r/w:t", nsm);

                // var anytext = tNodes.Any(tn => String.IsNullOrWhiteSpace(tn.Value));
                // $"TabCount  = {tabs.Count}\n<r> count = {rNodes.Count()}\n<t> count = {tNodes.Count()}\n".Dump($"Paragraph {paraNum}");

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
                        contentObj.AddTextObject(textNode.TabPos, yPos, textNode.Text, CourierNew, fontSize, textNode.Justification);
                    }
                }
                yPos += (int)(fontSize * vScale);
            }
            //NOT USED ANYMORE - Get previous TAB value
            //var previousTab = textNodes[1].XPathEvaluate("string(preceding-sibling::*[1]/w:tab/@pt14:TabWidth)",nsm).Dump("Previous Run TabWidth") ;
            //var tabPos = double.Parse(textNodes[1].PreviousNode.XPathSelectElement("w:tab",nsm).Attribute(pt14 + "TabWidth")?.Value) / 20;
            return pdf;
        }

        public void Write(string fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Create);

            int size = 0;
            file.Write(RenderHeader("1.4", out size));
            //Just do first page for now...
            file.Write(pageObjects[0].RenderPageRefs(file.Length, out size), 0, size);
            file.Write(pageObjects[0].RenderPageContent(file.Length, out size), 0, size);

            file.Write(catalogObj.RenderBytes(file.Length, out size), 0, size);
            file.Write(pageTreeObj.RenderBytes(file.Length, out size), 0, size);

            file.Write(pageObjects[0].RenderPageFonts(file.Length,out size),0,size);

            file.Write(infoObject.RenderBytes(file.Length, out size), 0, size);
            file.Write(xrefTable.RenderBytes(file.Length, out size), 0, size);
            file.Write(RenderTrailer(catalogObj.PdfObjectId, infoObject.PdfObjectId, out size), 0, size);
            file.Close();
        }
    }
}
