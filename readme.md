![logo](https://gitlab.3squared.com/uploads/-/system/project/avatar/470/pdf7.jpg)

# DocXToPdf

![image](https://gitlab.3squared.com/RobHill/docxtopdf/badges/debug/coverage.svg?style=flat-square)

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

## References used :
[PDF Reference](https://www.adobe.com/devnet/pdf/pdf_reference.html)
[O'Reilly Developing with PDF eBook](https://www.oreilly.com/library/view/developing-with-pdf/9781449327903/ch01.html)

[Minimal PDF file template](https://brendanzagaeski.appspot.com/0004.html)

[Hand coding PDF tutorial](https://brendanzagaeski.appspot.com/0005.html)

[Informal intro to DocX](https://www.toptal.com/xml/an-informal-introduction-to-docx)

[Embedding subset fonts into PDF](http://etutorials.org/Linux+systems/pdf+hacks/Chapter+4.+Creating+PDF+and+Other+Editions/Hack+43+Embed+and+Subset+Fonts+to+Your+Advantage/)

[WordProcessigML Spec (Paragraphs etc)](http://officeopenxml.com/WPparagraphProperties.php)


