![logo](https://gitlab.3squared.com/uploads/-/system/project/avatar/470/pdf7.jpg)

# DocXToPdf

| Release | Debug | Release | Debug |
|------|-------|---------|---------|
|![image](https://gitlab.3squared.com/RobHill/docxtopdf/badges/release/coverage.svg?style=flat-square)|![image](https://gitlab.3squared.com/RobHill/docxtopdf/badges/debug/coverage.svg?style=flat-square)|![image](http://teamcity.3squared.com/app/rest/builds/buildType:(id:DocXToPdf_Release)/statusIcon)|![image](http://teamcity.3squared.com/app/rest/builds/buildType:(id:DocXToPdf_Debug)/statusIcon)|




Converts simple Microsoft Word DocX documents to PDF files.

Currently only supports the first page (for the ORS implementation) and Monospaced fonts. Uses only a basic set of the PDF standard objects.

## Features

- **No 3rd party dependencies** - all plain vanilla C# 7.0
- **No GDI Dependencies** - (Main reason for monospaced fonts at the moment).
- **.Net Standard 2.0** - Plays well with .net & .net core projects.
- **Tested in Azure** - No gdi dependencies means it works in the cloud.

> NOTE: the document.xml is from within the MicrosoftWord .docx file. 
> This library _does not_ crack open the word document, it only converts the xml content file.

## Usage:
Read and write to a file:
```cs

var reader = XmlReader.Create(@"document.xml");
var xdoc = XDocument.Load(reader);
PdfDocument.FromDocX(xdoc).Write(@"output.pdf");

```
There is also an XDocument convenience extension method: 
```cs
var reader = XmlReader.Create(@"document.xml");
var xdoc = XDocument.Load(reader);
xdoc.ToPdf().Write(@"output.pdf");
```
Or returning a stream :
```cs
var memoryStream = new MemoryStream();
var pdf = PdfDocument.FromDocX(fileName);
var outputName = Path.ChangeExtension(fileName, "pdf");
pdf.Write(out memoryStream);
using (FileStream file = new FileStream(outputName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
{
    memoryStream.WriteTo(file);
    file.Close();
}
```

## References used :
[PDF Reference](https://www.adobe.com/devnet/pdf/pdf_reference.html)
[O'Reilly Developing with PDF eBook](https://www.oreilly.com/library/view/developing-with-pdf/9781449327903/ch01.html)

[Minimal PDF file template](https://brendanzagaeski.appspot.com/0004.html)

[Hand coding PDF tutorial](https://brendanzagaeski.appspot.com/0005.html)

[Informal intro to DocX](https://www.toptal.com/xml/an-informal-introduction-to-docx)

[Embedding subset fonts into PDF](http://etutorials.org/Linux+systems/pdf+hacks/Chapter+4.+Creating+PDF+and+Other+Editions/Hack+43+Embed+and+Subset+Fonts+to+Your+Advantage/)

[WordProcessigML Spec (Paragraphs etc)](http://officeopenxml.com/WPparagraphProperties.php)

[Portable Document Format: An Introduction for Programmers](http://preserve.mactech.com/articles/mactech/Vol.15/15.09/PDFIntro/index.html)

[Points, inches and Emus: Measuring units in Office Open XML](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/)