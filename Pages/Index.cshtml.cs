using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Hosting.Internal;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Word;
using word = Microsoft.Office.Interop.Word;

namespace wordtemplate2pdf.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly IHostingEnvironment _hst;

        public IndexModel(ILogger<IndexModel> logger, IHostingEnvironment hosting)
        {
            _logger = logger;
            _hst = hosting;
        }

        [BindProperty]
        public string FormFile { get; set; }

        public void OnGet()
        {

        }
      
        public void OnPost()
        {
            string path = _hst.WebRootPath + "/Template";
            Directory.CreateDirectory(path);

            string pathWithFileName = path + "/abc.docx";

            string tempFilePath = CloneTemplateForEditing(pathWithFileName);

            Application app = new word.Application();
            Document doc = app.Documents.Open(tempFilePath);

            
           FindAndReplace(app, "{{name}}", "Mr X");
           FindAndReplace(app, "{{email}}", "email@email.com");

            path = _hst.WebRootPath + "/TemplateParsed/"+ new Guid("hhghghghgh")+ ".pdf";
            Directory.CreateDirectory(_hst.WebRootPath + "/TemplateParsed/");
            doc.ExportAsFixedFormat(path, WdExportFormat.wdExportFormatPDF);

        }

        public string CloneTemplateForEditing(string templatePath)
        {
            var tempFile = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + Path.GetExtension(templatePath);
            System.IO.File.Copy(templatePath, tempFile);
            return tempFile;
        }

        public void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, string findText, string replaceWithText)
        {
            if (replaceWithText.Length > 255)
            {
                FindAndReplace(doc, findText, findText + replaceWithText.Substring(255));
                replaceWithText = replaceWithText.Substring(0, 255);
            }
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}
