using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;

namespace MyApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTable studentData = GetExcelData();
            var groupedData = from r in studentData.Rows.OfType<DataRow>()
                              group r by r["SCHOOL NAME"] into g
                              select new { School = g.Key, Data = g };
            string schoolName = "";

            foreach (var schoolData in groupedData)
            {
                if (!string.IsNullOrWhiteSpace(schoolData.School.ToString()))
                {
                    PdfDocument PDFDoc = PdfReader.Open(@"dest\VJS ADMITCARD.pdf", PdfDocumentOpenMode.Import);
                    PdfDocument PDFNewDoc = new PdfDocument();
                    foreach (DataRow row in schoolData.Data)
                    {
                        schoolName = row[5].ToString();
                        if (!string.IsNullOrWhiteSpace(row[0].ToString()))
                        {
                            PdfPage pp = PDFNewDoc.AddPage(PDFDoc.Pages[0]);
                            XGraphics gfx = XGraphics.FromPdfPage(pp);
                            XFont font = new XFont("Arial", 13, XFontStyle.Regular);
                            gfx.DrawString(row[7].ToString(), font, XBrushes.Black, new XRect(165, 122, 10, 10), XStringFormats.CenterLeft);
                            gfx.DrawString(row[1].ToString(), font, XBrushes.Black, new XRect(165, 142, 10, 10), XStringFormats.CenterLeft);
                            gfx.DrawString(row[2].ToString(), font, XBrushes.Black, new XRect(165, 162, 10, 10), XStringFormats.CenterLeft);
                            gfx.DrawString(row[6].ToString(), font, XBrushes.Black, new XRect(165, 183, 10, 10), XStringFormats.CenterLeft);
                            gfx.DrawString(row[4].ToString(), font, XBrushes.Black, new XRect(165, 202, 10, 10), XStringFormats.CenterLeft);
                            gfx.DrawString(row[5].ToString(), font, XBrushes.Black, new XRect(165, 222, 10, 10), XStringFormats.CenterLeft);
                        }
                    }
                    PDFNewDoc.Save(@"dest\AdmitCard_" + schoolName + ".pdf");
                }

            }
            Console.WriteLine("HallTickerts are Generated....Press any key to close the app");
            Console.ReadLine();
        }

        private static DataTable GetExcelData()
        {
            DataTable dt = new DataTable();
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=src\Students Master data.xlsx;Extended Properties='Excel 12.0;HDR=yes'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {

                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Sheet1$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    dt.Load(dr);
                    return dt;
                }
            }
        }
    }
}