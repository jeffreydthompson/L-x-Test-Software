using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PdfSharp.Drawing;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Microsoft.Win32;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using System.Windows;

namespace Luminex_Test_Software
{

    public delegate void PDFSaveComplete(bool savedSuccessfully);

    class PDFCreator
    {

        public PDFSaveComplete pdfSaveCompleteDelegate;

        public void SaveToPDF(string serialNumber, Dictionary<string, CablePin> results)
        {
            PdfDocument pdf = new PdfDocument();
            pdf.Info.Title = "Test Results";

            PdfPage page = pdf.AddPage();

            XGraphics gfx = XGraphics.FromPdfPage(page);

            // TITLE
            XFont font = new XFont("Arial", 30, XFontStyle.Regular);
            XStringFormat titleFormat = new XStringFormat
            {
                Alignment = XStringAlignment.Center
            };

            gfx.DrawString($"0110-8481 S/N {serialNumber} Test Results", font, XBrushes.CadetBlue, new XPoint( page.Width/2, 40), titleFormat);

            font = new XFont("Arial", 12, XFontStyle.Regular);
            XStringFormat timeStampFormat = new XStringFormat
            {
                Alignment = XStringAlignment.Far
            };
            gfx.DrawString($"{DateTime.Now}", font, XBrushes.CadetBlue, new XPoint(page.Width - 25, page.Height - 25), timeStampFormat);

            // Cable 4140
            double cable4140yStart = 120;
            font = new XFont("Arial", 20, XFontStyle.Regular);
            gfx.DrawString("Cable 3000-4140", font, XBrushes.DarkSlateGray, new XPoint(50, cable4140yStart));
            gfx.DrawLine(XPens.DarkSlateGray, 50, cable4140yStart+10, page.Width - 50, cable4140yStart+10);

            font = new XFont("Arial", 12, XFontStyle.Regular);
            gfx.DrawString("Pin Number", font, XBrushes.DarkSlateGray, new XPoint(60, cable4140yStart + 25));
            gfx.DrawString("Wire Color", font, XBrushes.DarkSlateGray, new XPoint(160, cable4140yStart + 25));
            gfx.DrawString("Measured Voltage", font, XBrushes.DarkSlateGray, new XPoint(260, cable4140yStart + 25));
            gfx.DrawString("Voltage Spec", font, XBrushes.DarkSlateGray, new XPoint(390, cable4140yStart + 25));
            gfx.DrawString("Pass / Fail", font, XBrushes.DarkSlateGray, new XPoint(490, cable4140yStart + 25));

            gfx.DrawLine(XPens.DarkSlateGray, 50, cable4140yStart + 28, page.Width - 50, cable4140yStart + 28);

            for (int i = 0; i < 5; i++)
            {
                double y = cable4140yStart + 46 + (i * 18);
                gfx.DrawLine(XPens.DarkSlateGray, 50, y, page.Width - 50, y);

                string key = $"_4140_{i+1}";

                gfx.DrawString($"{i+1}", font, XBrushes.DarkSlateGray, new XPoint(60, y-3));
                gfx.DrawString($"{results[key].WireColor}", font, XBrushes.DarkSlateGray, new XPoint(160, y-3));
                gfx.DrawString($"{results[key].MeasuredVoltage}V", font, XBrushes.DarkSlateGray, new XPoint(260, y-3));
                gfx.DrawString($"{results[key].VoltageRangeLow}V to {results[key].VoltageRangeHigh}V", 
                    font, XBrushes.DarkSlateGray, new XPoint(390, y-3));
                gfx.DrawString(results[key].passes() ? "Pass" : "Fail", 
                    font, XBrushes.DarkSlateGray, new XPoint(500, y-3));
            }

            // Cable 4130
            double cable4130yStart = 300;
            font = new XFont("Arial", 20, XFontStyle.Regular);
            gfx.DrawString("Cable 3000-4130", font, XBrushes.DarkSlateGray, new XPoint(50, cable4130yStart));
            gfx.DrawLine(XPens.DarkSlateGray, 50, cable4130yStart+10, page.Width - 50, cable4130yStart + 10);

            font = new XFont("Arial", 12, XFontStyle.Regular);
            gfx.DrawString("Pin Number", font, XBrushes.DarkSlateGray, new XPoint(60, cable4130yStart + 25));
            gfx.DrawString("Wire Color", font, XBrushes.DarkSlateGray, new XPoint(160, cable4130yStart + 25));
            gfx.DrawString("Measured Voltage", font, XBrushes.DarkSlateGray, new XPoint(260, cable4130yStart + 25));
            gfx.DrawString("Voltage Spec", font, XBrushes.DarkSlateGray, new XPoint(390, cable4130yStart + 25));
            gfx.DrawString("Pass / Fail", font, XBrushes.DarkSlateGray, new XPoint(490, cable4130yStart + 25));

            gfx.DrawLine(XPens.DarkSlateGray, 50, cable4130yStart + 28, page.Width - 50, cable4130yStart + 28);

            for (int j = 0; j < 14; j++)
            {
                double y = cable4130yStart + 46 + (j * 18);
                gfx.DrawLine(XPens.DarkSlateGray, 50, y, page.Width - 50, y);

                gfx.DrawString($"{j + 1}", font, XBrushes.DarkSlateGray, new XPoint(60, y - 3));

                if ((j+1) == 6 || (j + 1) > 12)
                {
                    gfx.DrawString($"Not Connected", font, XBrushes.DarkSlateGray, new XPoint(160, y - 3));
                } 
                else
                {
                    string key = $"_4130_{j + 1}";

                    gfx.DrawString($"{results[key].WireColor}", font, XBrushes.DarkSlateGray, new XPoint(160, y - 3));
                    gfx.DrawString($"{results[key].MeasuredVoltage}V", font, XBrushes.DarkSlateGray, new XPoint(260, y - 3));
                    gfx.DrawString($"{results[key].VoltageRangeLow}V to {results[key].VoltageRangeHigh}V",
                        font, XBrushes.DarkSlateGray, new XPoint(390, y - 3));
                    gfx.DrawString(results[key].passes() ? "Pass" : "Fail",
                        font, XBrushes.DarkSlateGray, new XPoint(500, y - 3));
                }
            }

            //XStringFormat timeStampFormat = new XStringFormat {
            //    Alignment = XStringAlignment.Far
            //};

            //string timeStamp = $"Tested at {DateTime.Now}";

            //gfx.DrawString(timeStamp, font, XBrushes.AliceBlue, new XPoint(800, 500), timeStampFormat);

            string filename = "";

            if (filename.Equals(""))
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "PDF | *.pdf";
                saveFile.Title = "Choose Save File Path";
                saveFile.ShowDialog();

                filename = saveFile.FileName;

            }

            try
            {
                pdf.Save(filename);
                //Process.Start(filename);
                pdfSaveCompleteDelegate(true);
            }
            catch (Exception e)
            {
                Console.Out.WriteLine($"Error: {e.Message}");
                MessageBox.Show($"Error.  Couldn't save file. {e.Message}");
                pdfSaveCompleteDelegate(false);
            }
        }

        public void HelloWorldPDF()
        {
            PdfDocument pdf = new PdfDocument();
            pdf.Info.Title = "Test pdf page";

            PdfPage page = pdf.AddPage();

            XGraphics gfx = XGraphics.FromPdfPage(page);

            XFont font = new XFont("Arial", 20, XFontStyle.Regular);

            //gfx.DrawString(
            //    "Hello, World!",
            //    font, 
            //    XBrushes.Black, 
            //    new XRect(0, 0, page.Width, page.Height), 
            //    XStringFormats.Center
            //);

            XPen pen = XPens.DarkGray;
            
            
            gfx.DrawLine(pen, 50, 50, page.Width - 50, 50);
            gfx.DrawString("testing", font, XBrushes.Black, new XPoint(100, 70));
            gfx.DrawLine(pen, 50, 80, page.Width - 50, 80);
            gfx.DrawString("testingPass", font, XBrushes.DarkBlue, new XPoint(100, 200));
            gfx.DrawString("testingFail", font, XBrushes.DarkRed, new XPoint(100, 300));

            string filename = "";

            if (filename.Equals(""))
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "PDF | *.pdf";
                saveFile.Title = "Choose Save File Path";
                saveFile.ShowDialog();

                filename = saveFile.FileName;
            }

            try
            {
                pdf.Save(filename);
            }
            catch (Exception e)
            {
                Console.Out.WriteLine($"Error: {e.Message}");
            }
        }
    }
}
