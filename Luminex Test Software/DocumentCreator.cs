using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace Luminex_Test_Software
{

    public delegate void DocumentSaveComplete(bool savedSuccessfully);

    class DocumentCreator
    {

        public DocumentSaveComplete documentSavedDelegateCallback;

        public void SaveResultsToFile(string serialNumber, Dictionary<string, CablePin> results)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Add();

            foreach (Section section in doc.Sections) // only one but foreach is safe & clean access
            {
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                headerRange.Font.Size = 26;
                headerRange.Text = "Luminex 0110-8481 Test Results";
            }

            doc.Content.SetRange(0, 0);

            Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Range.set_Style("Heading 1");
            paragraph.Range.Text = $"Serial Number: {serialNumber}";
            paragraph.Range.InsertParagraphAfter();

            Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Range.set_Style("Heading 1");
            paragraph2.Range.Text = "Cable 3000-4130";
            paragraph2.Range.InsertParagraphAfter();

            Table table = doc.Tables.Add(paragraph2.Range, 15, 5);

            table.Borders.Enable = 1;
            //table.Rows.Height = 12;

            foreach (Row row in table.Rows)
            {
                if (row.Index == 1)
                {
                    row.Cells[1].Range.Text = "Pin Number";
                    row.Cells[2].Range.Text = "Wire Color";
                    row.Cells[3].Range.Text = "Measured Voltage";
                    row.Cells[4].Range.Text = "Voltage Spec";
                    row.Cells[5].Range.Text = "Pass / Fail";

                    foreach (Cell cell in row.Cells)
                    {
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }
                }
                else if (row.Index == 7 || row.Index == 14 || row.Index == 15) // +1 because title rows is index 1
                {
                    foreach (Cell cell in row.Cells)
                    {
                        if (cell.ColumnIndex == 1)
                        {
                            cell.Range.Text = $"{row.Index - 1}";
                        }
                        else if (cell.ColumnIndex == 2)
                        {
                            cell.Range.Text = "Not Connected";
                        }
                        else
                        {
                            cell.Range.Text = "-";
                            cell.Shading.BackgroundPatternColor = WdColor.wdColorGray75;
                        }
                    }
                }
                else
                {
                    string key = $"_4130_{row.Index - 1}";

                    row.Cells[1].Range.Text = $"{row.Index - 1}";
                    row.Cells[2].Range.Text = results[key].WireColor;
                    row.Cells[3].Range.Text = $"{results[key].MeasuredVoltage}V";
                    row.Cells[4].Range.Text = $"{results[key].VoltageRangeLow}V to {results[key].VoltageRangeHigh}V";
                    row.Cells[5].Range.Text = results[key].passes() ? "Pass" : "Fail";
                }
            }

            Paragraph paragraph3 = doc.Content.Paragraphs.Add();
            paragraph3.Range.set_Style("Heading 1");
            paragraph3.Range.Text = "Cable 3000-4140";
            paragraph3.Range.InsertParagraphAfter();

            Table table2 = doc.Tables.Add(paragraph3.Range, 6, 5);

            table2.Borders.Enable = 1;
            //table2.Rows.Height = 12;

            //Console.Out.WriteLine($"table2 row count: {table2.Rows.Count}");

            foreach (Row row in table2.Rows)
            {
                if (row.Index == 1)
                {
                    row.Cells[1].Range.Text = "Pin Number";
                    row.Cells[2].Range.Text = "Wire Color";
                    row.Cells[3].Range.Text = "Measured Voltage";
                    row.Cells[4].Range.Text = "Voltage Spec";
                    row.Cells[5].Range.Text = "Pass / Fail";

                    foreach (Cell cell in row.Cells)
                    {
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }
                }
                else
                {
                    string key = $"_4140_{row.Index - 1}";

                    //Console.Out.WriteLine($"row index: {row.Index}, key: {key}");

                    row.Cells[1].Range.Text = $"{row.Index - 1}";
                    row.Cells[2].Range.Text = results[key].WireColor;
                    row.Cells[3].Range.Text = $"{results[key].MeasuredVoltage}V";
                    row.Cells[4].Range.Text = $"{results[key].VoltageRangeLow}V to {results[key].VoltageRangeHigh}V";
                    row.Cells[5].Range.Text = results[key].passes() ? "Pass" : "Fail";
                }
            }

            try
            {
                //object filename = @"C:\Users\Jeffrey\Documents\learning\Temp\temp1.docx";

                object filename = "";

                if (filename.Equals(""))
                {
                    SaveFileDialog saveFile = new SaveFileDialog();
                    saveFile.Filter = "Word Document | *.docx";
                    saveFile.Title = "Choose Save File Path";
                    saveFile.ShowDialog();

                    filename = saveFile.FileName;
                }


                doc.SaveAs(filename);
                doc.Close();
                app.Quit();
                MessageBox.Show("Saved test file successfully");
                documentSavedDelegateCallback(true);
            }
            catch (Exception error)
            {
                documentSavedDelegateCallback(false);
                Console.Out.WriteLine($"Error: {error.Message}");
            }
        }

        public void TestDocument()
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Add();

            Console.Out.WriteLine($"doc sections count: {doc.Sections.Count}");

            foreach (Section section in doc.Sections) // only one but foreach is safe & clean access
            {
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                headerRange.Font.Size = 12;
                headerRange.Text = "This is a header";

                Range footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                footerRange.Font.ColorIndex = WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 10;
                footerRange.Text = "This is a footer";
            }

            doc.Content.SetRange(0, 0);
            doc.Content.Text = "This is the documents text";

            Paragraph paragraph = doc.Content.Paragraphs.Add();
            paragraph.Range.set_Style("Heading 1");
            paragraph.Range.Text = $"Serial Number: {5468521325}";
            paragraph.Range.InsertParagraphAfter();

            Paragraph paragraph2 = doc.Content.Paragraphs.Add();
            paragraph2.Range.set_Style("Heading 1");
            paragraph2.Range.Text = "Cable 3000-4130";
            paragraph2.Range.InsertParagraphAfter();

            Table table = doc.Tables.Add(paragraph2.Range, 5, 4);

            table.Borders.Enable = 1;

            foreach (Row row in table.Rows)
            {
                if (row.Index == 1)
                {
                    row.Cells[1].Range.Text = "Pin Number";
                    row.Cells[2].Range.Text = "Measured Voltage";
                    row.Cells[3].Range.Text = "Voltage Specification";
                    row.Cells[4].Range.Text = "Pass / Fail";

                    foreach (Cell cell in row.Cells)
                    {
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                    }
                }
                else
                {

                }
            }

            try
            {
                //object filename = @"C:\Users\Jeffrey\Documents\learning\Temp\temp1.docx";

                object filename = "";

                if (filename.Equals(""))
                {
                    SaveFileDialog saveFile = new SaveFileDialog();
                    saveFile.Filter = "Word Document | *.docx";
                    saveFile.Title = "Choose Save File Path";
                    saveFile.ShowDialog();

                    filename = saveFile.FileName;
                }


                doc.SaveAs(filename);
                doc.Close();
                app.Quit();
                MessageBox.Show("Saved test file successfully");
            }
            catch (Exception error)
            {
                Console.Out.WriteLine($"Error: {error.Message}");
            }

        }

    }
}
