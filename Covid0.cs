
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;



namespace CovidVoucheri
{
    public static class Covid0
    {

        static void MergePDFs(string targetPath, List<String> pdfs)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (PdfDocument targetDoc = new PdfDocument())
            {
                foreach (string pdf in pdfs)
                {
                    using (PdfDocument pdfDoc = PdfReader.Open(pdf, PdfDocumentOpenMode.Import))
                    {
                        for (int i = 0; i < pdfDoc.PageCount; i++)
                        {
                            targetDoc.AddPage(pdfDoc.Pages[i]);
                        }
                    }
                }
                targetDoc.Save(targetPath+=@"\FINAL.pdf");
            }
        }
        static string AddSuffix(string filename, string suffix)
        {
            string fDir = Path.GetDirectoryName(filename);
            string fName = Path.GetFileNameWithoutExtension(filename);
            string fExt = Path.GetExtension(filename);
            return Path.Combine(fDir, String.Concat(fName, suffix, fExt));
        }
        static void FindAndReplace(Microsoft.Office.Interop.Word.Application app, Document doc, string findText, string replaceWithText)
        {
            Find findObject = app.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = findText;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replaceWithText;
            object missing = System.Reflection.Missing.Value;
            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceAll, ref missing, ref missing, ref missing, ref missing);

            var shapes = doc.Shapes;
            foreach (Microsoft.Office.Interop.Word.Shape shape in shapes)
            {
                if (shape.TextFrame.HasText != 0)
                {
                    var initialText = shape.TextFrame.TextRange.Text;
                    var resultingText = initialText.Replace(findText, replaceWithText);
                    if (initialText != resultingText)
                    {
                        shape.TextFrame.TextRange.Text = resultingText;
                    }
                }
            }
        }
        static void CreateWordDocument(object filename, object SaveAs, string ime, string prezime, string datum, string email, string broj)
        {

            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;

            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                FindAndReplace(wordApp, myWordDoc, "<<ime>>", ime);
                FindAndReplace(wordApp, myWordDoc, "<<prezime>>", prezime);
                FindAndReplace(wordApp, myWordDoc, "<<datum>>", datum);
                FindAndReplace(wordApp, myWordDoc, "<<email>>", email);
                FindAndReplace(wordApp, myWordDoc, "<<broj>>", broj);
            }
            else
            {
                Console.WriteLine("error");
            }
            try
            {
                object fileFormat = Word.WdSaveFormat.wdFormatPDF;

                myWordDoc.SaveAs(ref SaveAs,
                                ref fileFormat, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing,
                                ref missing, ref missing, ref missing, ref missing);
            }


            catch (Exception e) { Console.WriteLine(e); }


            myWordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            wordApp.Quit();

        }
        public static void Iscitaj(int xlOd, int xlDo, int Sheet, string path, string filename)
        {
            List<string> files = new List<string>();
            Microsoft.Office.Interop.Excel._Application excel = new _Excel.Application();
            Workbook wb;
            Worksheet ws;



            wb = excel.Workbooks.Open(@path);
            ws = wb.Worksheets[Sheet];



            
            string ime, prezime, datum, email, broj;


            int n = 0;
            int x = xlOd;
            int j = xlDo;

            var voucher_path = Path.Combine(Directory.GetCurrentDirectory(), @"\voucher.docx");
            for (int i = x; i < j; i++)
            {

                if (ws.Cells[i, 1].Value2 != null)
                {
                    ime = ws.Cells[i, 1].Value2;
                    prezime = ws.Cells[i, 2].Value2;
                    datum = ws.Cells[i, 3].Value2;
                    email = ws.Cells[i, 5].Value2;
                    broj = ws.Cells[i, 4].Value2.ToString();



                    //napravi word, spremi
                    string newFilename = AddSuffix(filename+@"\voucher.pdf", String.Format("{0}", n));
                   
                 
                    CreateWordDocument(voucher_path, newFilename, ime, prezime, datum, email, broj);
                    files.Add(newFilename);
          

                    n++;




                }
                MergePDFs(filename, files);
            }



            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
    
          




    }
}
