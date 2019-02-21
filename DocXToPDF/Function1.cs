
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
    public static class Function1
    {
        [FunctionName("docx2pdf")]
        public static IActionResult Run([HttpTrigger(AuthorizationLevel.Anonymous,"get", "post", Route = null)]HttpRequest req, TraceWriter log, ExecutionContext context)
        {
            log.Info("DocX2Pdf function request.");
            //Stopwatch sw = new Stopwatch();
            string fileName = req.Query["name"];
            string requestBody = new StreamReader(req.Body).ReadToEnd();
            //sw.Start();
            XDocument xdoc = null;
            if (string.IsNullOrEmpty(requestBody))
            {
                string samplFileName = Path.Combine(context.FunctionAppDirectory, "Data", "file_1.xml");
                xdoc = XDocument.Load(samplFileName);
                fileName = "SampleViaHttpGet.pdf";
            }
            else
            {
                xdoc = XDocument.Parse(requestBody);
            }
            var pdfDoc = xdoc.ToPdf();
            var str = pdfDoc.GetBytes(new MemoryStream());
            var result = new FileContentResult(str, "application/pdf");
            if (!string.IsNullOrEmpty(fileName))
                fileName = Path.ChangeExtension(fileName, "pdf");
            result.FileDownloadName = fileName ?? "download.pdf";
            //sw.Stop();
            //log.Info($"Word DocX to Pdf convertion time : {sw.ElapsedMilliseconds}ms");
            return result;
        }
    }
}
