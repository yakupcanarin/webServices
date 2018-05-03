using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Services;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace webApplication
{
    /// <summary>
    /// Summary description for DBEntegration
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class DBEntegration : System.Web.Services.WebService
    {
        static string connString = ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
        SqlConnection sqlConn = new SqlConnection(connString);

        [WebMethod]
        public byte[] CreateArchiveFile(byte[] file, string docType)
        {
            Random rnd = new Random();
            int number = rnd.Next(111111);
            string path = @"C:\Zip\"+number+"."+docType;
            string zippedPath = @"C:\Zip\Zipped\" + number + ".zip";
            File.WriteAllBytes(path, file);


            byte[] read = new byte[4096];
            int readByte = 0;
            MemoryStream _mStream = new MemoryStream(); // create a empty memory
            ZipArchive archive = new ZipArchive(_mStream, ZipArchiveMode.Create, true); // create zip archive into the memory
            ZipArchiveEntry fileArchive = archive.CreateEntry(path); // show the document to archive which is in memory
            var OpenFileinArchive = fileArchive.Open(); //create document in memory archive
            FileStream _fsReader = new FileStream(path, FileMode.Open, FileAccess.Read); //get file set permissions.

            while (_fsReader.Position != _fsReader.Length)
            {
                readByte = _fsReader.Read(read, 0, read.Length);  // read file
                OpenFileinArchive.Write(read, 0, readByte);  // write to memory
            }

            _fsReader.Dispose();
            OpenFileinArchive.Close();
            archive.Dispose();

            using (var _fs = new FileStream(zippedPath, FileMode.Create))
            {
                _mStream.Seek(0, SeekOrigin.Begin);
                _mStream.CopyTo(_fs);
            }

            return File.ReadAllBytes(zippedPath);


        }
        [WebMethod]
        public byte[] ConvertToPDF(byte[] data, string extention)
        {
            var dataDosya = @"C:\Document";
            var pdfDosya = @"C:\PDF";
            Random rnd = new Random();
            int randomSayi = rnd.Next(111111);
            string dosyaPathdata = dataDosya + @"\" + randomSayi + "."+extention;
            string dosyapathPdf = pdfDosya + @"\" + randomSayi + ".pdf";
            File.WriteAllBytes(dosyaPathdata, data);
            if (extention == "xlsx")
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                app.Visible = false;
                Microsoft.Office.Interop.Excel.Workbook wkb = app.Workbooks.Open(dosyaPathdata, ReadOnly: true);
                wkb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, dosyapathPdf);
                wkb.Close();
                app.Quit();
            }
            else if (extention == "doc" || extention == "docx")
            {
                Microsoft.Office.Interop.Word.Application wApp = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document wordDoc = null;
                object inputFileTemp = dosyaPathdata;
                wordDoc = wApp.Documents.Open(dosyaPathdata);
                wordDoc.ExportAsFixedFormat(dosyapathPdf, WdExportFormat.wdExportFormatPDF);
                wordDoc.Close();
                wApp.Quit();
            }
            
            var pdfByte = File.ReadAllBytes(pdfDosya + @"\" + randomSayi + ".pdf");
            if (File.Exists(dosyaPathdata))
            {
                try
                {
                    File.Delete(dosyaPathdata);
                }
                catch (Exception)
                {
                    throw;
                }
            }
            //if (File.Exists(dosyapathPdf))
            //{
            //    try
            //    {
            //        File.Delete(dosyapathPdf);
            //    }
            //    catch (Exception)
            //    {
            //        throw;
            //    }
            //}
            return pdfByte;

        }
    }
}
