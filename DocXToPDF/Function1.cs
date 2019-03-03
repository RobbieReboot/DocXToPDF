
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocxToPdf.Core;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
//using DocxToPdf.Core;
namespace DocXToPDF
{
    public static class DocXtoPdfAzureFunction
    {
        [FunctionName("docx2pdf")]
        public static IActionResult Run([HttpTrigger(AuthorizationLevel.Anonymous,"get", "post", Route = null)]HttpRequest req, TraceWriter log, ExecutionContext context)
        {
            log.Info("DocX2Pdf function request.");
            //Stopwatch sw = new Stopwatch();
            string fileName = req.Query["name"];

            var memoryStream = new MemoryStream();
            req.Body.CopyTo(memoryStream);

            XDocument xdoc = null;
            PdfDocument pdfDoc;

            if (req.Body.Length==0)
            {
                return new NoContentResult();
            }
            else
            {
                pdfDoc = PdfDocument.FromDocX(memoryStream);
            }

            var str = pdfDoc.GetBytes(new MemoryStream());
            var result = new FileContentResult(str, "application/pdf");
            if (!string.IsNullOrEmpty(fileName))
                fileName = Path.ChangeExtension(fileName, "pdf");
            result.FileDownloadName = fileName ?? "download.pdf";
            return result;
        }
    }
}
