using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace ajaxtest3
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        [System.Web.Services.WebMethod]
        public static string DownloadDoc(string fileName)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            //WordprocessingDocument wordprocessingDocument =
            //    WordprocessingDocument.Open("C:\\Users\\admin\\Desktop\\tmpeng\\test.docx", true);
            //Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

            //Paragraph para = body.AppendChild(new Paragraph());
            //Run run = para.AppendChild(new Run());
            //run.AppendChild(new Text("ДОБАВЛЕНО"));

            //// Close the handle explicitly.
            //wordprocessingDocument.Close();

            byte[] byteArray = File.ReadAllBytes("C:\\Users\\admin\\Desktop\\tmpeng\\test.docx");
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;

                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("ДОБАВЛЕНО222222222222"));      // Do work here
                    body.InnerXml = body.InnerXml.Replace("{fff}", "ЗАМЕНЕНО");
                }
                // Save the file with the new name
                var bytetoreturn = stream.ToArray();
                return Convert.ToBase64String(bytetoreturn, 0, bytetoreturn.Length);
            }



            return "";

        }
        public void  trydoc()
        {
            byte[] byteArray = File.ReadAllBytes("c:\\data\\hello.docx");
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    // Do work here
                }
                // Save the file with the new name
                File.WriteAllBytes("C:\\data\\newFileName.docx", stream.ToArray());
            }
        }

        [System.Web.Services.WebMethod]
        public static string DownloadFile(string fileName)
        {
            //Set the File Folder Path.
            //string path = HttpContext.Current.Server.MapPath("~/Files/");

            //Read the File as Byte Array.
            byte[] bytes = File.ReadAllBytes("C:\\Users\\admin\\Desktop\\tmpeng\\test.docx");

            //Convert File to Base64 string and send to Client.
            return Convert.ToBase64String(bytes, 0, bytes.Length);
        }
    }
}