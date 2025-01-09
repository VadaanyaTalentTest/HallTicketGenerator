using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using ExcelDataReader;

 
namespace MyApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        public static double height = 3;
        public static double width = 3;
        public static int NumHallTicketPage = 2;

        public static int[,,] PageCoord = new int[3, 9, 2] { { { 125, 130 }, { 135, 153 }, { 410, 153 }, { 410,105},
                                                                {410,115},{410,130},{100,105},{96,213}, { 115, 105} },
                                                              { { 125, 413 }, { 135, 436 }, { 410, 436 }, { 410,388},
                                                                {410,398},{410,413},{100,388},{96,496}, { 115, 388} },
                                                               { { 125, 696 }, { 135, 719 }, { 410, 719 }, { 410,671},
                                                                {410,681},{410,696},{100,671},{96,779}, { 115, 671}}};
        static void Main(string[] args)
        {
            // System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTable studentData = null;
            try
            {
                studentData = GetExcelData();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error using OLEDB: " + ex.Message);
                studentData = GetExcelData2();
            }
            var groupedData = from r in studentData.Rows.OfType<System.Data.DataRow>()
                              group r by r["SCHOOL"] into g
                              select new { School = g.Key, Data = g };
            string schoolName = "";
            string MandalName = "";
            string DestinationPath = @"..\..\dest\";
            string MandalSpecificDestinationPath = "";
            int HallTicketIndex = 0;

            foreach (var schoolData in groupedData)
            {
                HallTicketIndex = 0;
                if (!string.IsNullOrWhiteSpace(schoolData.School.ToString()))
                {
                    PdfDocument PDFDoc = PdfReader.Open(@"dest\HALLTCKT2024-V2.pdf", PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
                    PdfDocument PDFNewDoc = new PdfDocument();
                    PdfPage pp = PDFNewDoc.AddPage(PDFDoc.Pages[0]);
                    XGraphics gfx = XGraphics.FromPdfPage(pp);
                    XFont font = new XFont("Arial", 10.5, XFontStyleEx.Bold);
                    XFont font2 = new XFont("Arial", 9.5, XFontStyleEx.Bold);
                    XFont font3 = new XFont("Arial", 7.5, XFontStyleEx.Bold);

                    foreach (DataRow row in schoolData.Data)
                    {
                        if(HallTicketIndex>NumHallTicketPage)
                        {
                            HallTicketIndex = 0;
                            pp = PDFNewDoc.AddPage(PDFDoc.Pages[0]);
                            gfx = XGraphics.FromPdfPage(pp);
                        }
                        schoolName = row[7].ToString();
                        MandalName = row[6].ToString();
                        if (!string.IsNullOrWhiteSpace(row[0].ToString()))
                        {
                            //Hallticket
							gfx.DrawString(row[12].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 8, 0], PageCoord[HallTicketIndex, 8, 1], height, width), XStringFormats.CenterLeft);
                            //Class
                            gfx.DrawString(row[3].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 3, 0], PageCoord[HallTicketIndex, 3, 1], height, width), XStringFormats.CenterLeft);
							//StudentName
							gfx.DrawString(row[0].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 0, 0], PageCoord[HallTicketIndex, 0, 1], height, width), XStringFormats.CenterLeft);
							//Exam centre
							gfx.DrawString(row[11].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 6, 0], PageCoord[HallTicketIndex, 7, 1], height, width), XStringFormats.CenterLeft);
							//Mobile Number
							gfx.DrawString(row[5].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 5, 0], PageCoord[HallTicketIndex, 5, 1], height, width), XStringFormats.CenterLeft);
							//Parent/Guardian Name
							gfx.DrawString(row[1].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 1, 0], PageCoord[HallTicketIndex, 1, 1], height, width), XStringFormats.CenterLeft);
							//School Name
							gfx.DrawString(row[7].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 2, 0], PageCoord[HallTicketIndex, 2, 1], height, width), XStringFormats.CenterLeft);


							//gfx.DrawString(row[7].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex,0,0], PageCoord[HallTicketIndex, 0, 1], height, width), XStringFormats.CenterLeft);
							//gfx.DrawString(row[1].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 1, 0], PageCoord[HallTicketIndex, 1, 1], height, width), XStringFormats.CenterLeft);
							//gfx.DrawString(row[2].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex,2, 0], PageCoord[HallTicketIndex, 2, 1], height, width), XStringFormats.CenterLeft);
                            //gfx.DrawString(row[4].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 5, 0], PageCoord[HallTicketIndex, 5, 1], height, width), XStringFormats.CenterLeft);
                            //gfx.DrawString(row[5].ToString(), font, XBrushes.Black, new XRect(PageCoord[HallTicketIndex, 6, 0], PageCoord[HallTicketIndex, 7, 1], height, width), XStringFormats.CenterLeft);
                        }
                        HallTicketIndex++;
                    }
                    MandalSpecificDestinationPath = DestinationPath + MandalName;
                    MandalSpecificDestinationPath = MandalSpecificDestinationPath.Replace(" ", "_");
                    if(!Directory.Exists(MandalSpecificDestinationPath))
                    {
                        Directory.CreateDirectory(MandalSpecificDestinationPath);
                    }
                    PDFNewDoc.Save(MandalSpecificDestinationPath+"\\AdmitCard_" + schoolName + ".pdf");
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

        private static DataTable GetExcelData2()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            DataTable dt = new DataTable();
            string filePath = @"..\..\..\src\Students Master data.xlsx";

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataReader.ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataReader.ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    dt = result.Tables[0];
                }
            }

            return dt;
        }
    }
}