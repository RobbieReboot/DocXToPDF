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
using System.IO.Compression;
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


        /*

public static void FromFullDocX(string filename)
{

	XDocument docx = null;
	XDocument styles = null;

	using (ZipArchive zip = ZipFile.Open(filename, ZipArchiveMode.Read))
	{
		var zdoc = zip.Entries.SingleOrDefault(n => n.FullName == @"word/document.xml");
		using (StreamReader s = new StreamReader(zdoc.Open()))
			docx = XDocument.Load(s);
		var zstyles = zip.Entries.SingleOrDefault(n => n.FullName == @"word/styles.xml");
		using (StreamReader s = new StreamReader(zstyles.Open()))
			styles = XDocument.Load(s);
	}
	var reader = docx.Root.CreateReader();
	var nsm = new XmlNamespaceManager(reader.NameTable);
	XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
	XNamespace pt14 = "http://powertools.codeplex.com/2011";
	nsm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
	nsm.AddNamespace("pt14", "http://powertools.codeplex.com/2011");

	var stylesWithTabs = styles.Descendants(w + "style")
		.Where(xe => xe.Attribute(w + "type")?.Value == "paragraph" &&
					 xe.Elements(w + "pPr").Elements(w + "tabs").Any())
		.ToDictionary(xe => xe.Attribute(w + "styleId").Value, xe => xe.Element(w + "pPr").Element(w + "tabs"));

	var paras = docx.Descendants(w + "p").ToList();
	var paraNum = 0;
	foreach (var para in paras)
	{
		$"Para({paraNum})".Dump();

		//Get the TABS table.
		//                var tabs = para.XPathSelectElements("w:pPr/w:tabs/w:tab", nsm).ToList();
		var rNodes = para.XPathSelectElements("w:r", nsm).ToList();
		var paraStyleName = para.Attributes(pt14 + "StyleName").FirstOrDefault()?.Value;
		var paraTabs = stylesWithTabs.FirstOrDefault(n => n.Key == paraStyleName);
		var tabTable = paraTabs.Value.Descendants(w + "tab").Select(t => new
		{
			Justification = t.Attribute(w + "val").Value,
			Position = t.Attribute(w + "pos").Value
		})
		.ToList();
		var currentTabInTable = 0;
		string textAccumulator = String.Empty;
		
		int nodeNum = 0;
		
		while (nodeNum < rNodes.Count())
		{
			bool consumeTab = false;
			string currentTab = String.Empty;
			//if the node has a <w:tab> USE the NEXT tab
			//if the node has text, accumulate this node (+ future nodes with <w:t> UNLESS another <w:tab> is present.
			if (rNodes[nodeNum].Descendants(w + "tab").Any())
			{
				if (currentTabInTable < tabTable.Count())
				{
					//break; //?
					currentTab = tabTable[currentTabInTable].Position;
					currentTabInTable++;
				}
				nodeNum++;          //Consume the tab.
				if (nodeNum >= rNodes.Count())
					break;
			}
			while (rNodes[nodeNum].Descendants(w+"t").Any())
			{
				textAccumulator += rNodes[nodeNum].Value;
				nodeNum++;
				if (nodeNum>=rNodes.Count())
					break;
			}
			$"{nodeNum} : {currentTab} : {textAccumulator}".Dump();
				textAccumulator = String.Empty;
		}


paraNum++;

	}
	return;
}

         
         */

        public static PdfDocument FromFullDocX(string filename)
        {

            XDocument docx = null;
            XDocument styles = null;

            using (ZipArchive zip = ZipFile.Open(filename, ZipArchiveMode.Read))
            {
                var zdoc = zip.Entries.SingleOrDefault(n => n.FullName == @"word/document.xml");
                using (StreamReader s = new StreamReader(zdoc.Open()))
                    docx = XDocument.Load(s);
                var zstyles = zip.Entries.SingleOrDefault(n => n.FullName == @"word/styles.xml");
                using (StreamReader s = new StreamReader(zstyles.Open()))
                    styles = XDocument.Load(s);
            }
            var reader = docx.Root.CreateReader();
            var nsm = new XmlNamespaceManager(reader.NameTable);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace pt14 = "http://powertools.codeplex.com/2011";
            nsm.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            nsm.AddNamespace("pt14", "http://powertools.codeplex.com/2011");

            var stylesWithTabs = styles.Descendants(w + "style")
                .Where(xe => xe.Attribute(w + "type")?.Value == "paragraph" &&
                             xe.Elements(w + "pPr").Elements(w + "tabs").Any())
                .ToDictionary(xe => xe.Attribute(w + "styleId").Value, xe => xe.Element(w + "pPr").Element(w + "tabs"));

            var paras = docx.Descendants(w + "p").ToList();
            foreach (var para in paras)
            {
                //Get the TABS table.
//                var tabs = para.XPathSelectElements("w:pPr/w:tabs/w:tab", nsm).ToList();
                var rNodes = para.XPathSelectElements("w:r", nsm).ToList();
                var paraStyleName = para.Attributes(pt14 + "StyleName").FirstOrDefault()?.Value;
                var paraTabs  = stylesWithTabs.FirstOrDefault(n => n.Key==paraStyleName);

                foreach (var rNode in rNodes)
                {
                    //if the node 
                }




            }

            return null;
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
            FontObject CourierNew = new FontObject("Courier",pdf);
            // TODO: Should come from the Docs metadata..
            pdf.AddInfoObject("XDoc2Pdf", "RobHill","3Squared");

            // TODO: Point sizes from adobe for A4 - Fixed for now. Margins fixed for now.
            var page = pdf.AddPage(new PageExtents(612, 792, 32, 10, 32, 10));
            // TODO: The text object should add the font to the page IF it's not already added (Possibly). Slower than this but more convenient.

            page.AddFont(CourierNew);

            // TODO: SHould refactor this so the PAGE object is responsible for creation of content object so the parentRef can be set on creation.
            var contentObj = page.AddContentObject();

            foreach (var para in paras)
            {
                //Get the TABS table.

                var tabs = para.XPathSelectElements("w:pPr/w:tabs/w:tab", nsm).ToList();
                var rNodes = para.XPathSelectElements("w:r", nsm).ToList();
                var tNodes = para.XPathSelectElements("w:r/w:t", nsm);

                // var anytext = tNodes.Any(tn => String.IsNullOrWhiteSpace(tn.Value));
                // $"TabCount  = {tabs.Count}\n<r> count = {rNodes.Count()}\n<t> count = {tNodes.Count()}\n".Dump($"Paragraph {paraNum}");

                var matchedTabNodes = new List<Tuple<string, double, string>>()
                          .Select(t => new { Text = string.Empty, TabPos = double.MinValue, Justification = String.Empty }).ToList();
                var tabNo = 0;
                var tabCount = tabs.Count();
                foreach (var rnode in rNodes)
                {
                    double tabStop = 0;
                    string justification = "left";

                    //Only create a tabstop IF we have tabs.
                    if (tabNo < tabs.Count)
                    {
                        tabStop = (double.Parse(tabs[tabNo].Attribute(w + "pos").Value) / 20) * hScale;
                        justification = tabs[tabNo].Attribute(w + "val").Value;
                    }

                    var newNode = new
                    {
                        Text = SanitizePdfCharacters(rnode.XPathSelectElement("w:t", nsm)?.Value),
                        TabPos = tabStop,  //(fontSize*1.6),
                        Justification = justification
                    };
                    matchedTabNodes.Add(newNode);
                        //if (!string.IsNullOrEmpty(newNode.Text))
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

        public static string SanitizePdfCharacters(string value)
        {
            if (String.IsNullOrEmpty(value))
                return String.Empty;

            var result = value.Replace(@"\", @"\\");    //MUST do the single slash first to avoid replacing the slashed in the FOLLOWING replacements!
            result = result.Replace(")", @"\)");
            result = result.Replace("(", @"\(");
            return result;
        }

        public byte[] GetBytes(MemoryStream ms)
        {
            int size = 0;
            ms.Write(RenderHeader("1.4", out size), 0, size);

            //Just do first page for now...
            ms.Write(pageObjects[0].RenderPageRefs(ms.Length, out size), 0, size);
            ms.Write(pageObjects[0].RenderPageContent(ms.Length, out size), 0, size);

            ms.Write(catalogObj.RenderBytes(ms.Length, out size), 0, size);
            ms.Write(pageTreeObj.RenderBytes(ms.Length, out size), 0, size);

            ms.Write(pageObjects[0].RenderPageFonts(ms.Length, out size), 0, size);

            ms.Write(infoObject.RenderBytes(ms.Length, out size), 0, size);
            ms.Write(xrefTable.RenderBytes(ms.Length, out size), 0, size);
            ms.Write(RenderTrailer(catalogObj.PdfObjectId, infoObject.PdfObjectId, out size), 0, size);
            ms.Close();
            return ms.GetBuffer();
        }

        public void Write(string fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Create);

            int size = 0;
            file.Write(RenderHeader("1.4", out size),0,size);

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
