using Spire.Xls;
using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core;

namespace ConditionalFormatRuntime
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            // Load file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ConditionalFormatRuntime.xlsx");

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];
            AddComparisonRule1(sheet);
            AddComparisonRule2(sheet);
            AddComparisonRule3(sheet);
            AddComparisonRule4(sheet);

            //Save to file
            String result = "ConditionalFormatRuntime_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer(result);
        }
        private void AddComparisonRule1(Worksheet sheet)
        {
            //Create conditional formatting rule
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range["A1:D1"]);
            IConditionalFormat cf1 = xcfs1.AddCondition();
            cf1.FormatType = ConditionalFormatType.CellValue;
            cf1.FirstFormula = "150";
            cf1.Operator = ComparisonOperatorType.Greater;
            cf1.FontColor = Color.Red;
            cf1.BackColor = Color.LightBlue;
        }
        private void AddComparisonRule2(Worksheet sheet)
        {
            XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
            xcfs2.AddRange(sheet.Range["A2:D2"]);
            IConditionalFormat cf2 = xcfs2.AddCondition();
            cf2.FormatType = ConditionalFormatType.CellValue;
            cf2.FirstFormula = "500";
            cf2.Operator = ComparisonOperatorType.Less;
            //Set border color
            cf2.LeftBorderColor = Color.Pink;
            cf2.RightBorderColor = Color.Pink;
            cf2.TopBorderColor = Color.DeepSkyBlue;
            cf2.BottomBorderColor = Color.DeepSkyBlue;
            cf2.LeftBorderStyle = LineStyleType.Medium;
            cf2.RightBorderStyle = LineStyleType.Thick;
            cf2.TopBorderStyle = LineStyleType.Double;
            cf2.BottomBorderStyle = LineStyleType.Double;
        }

        private void AddComparisonRule3(Worksheet sheet)
        {
            //Create conditional formatting rule
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range["A3:D3"]);
            IConditionalFormat cf1 = xcfs1.AddCondition();
            cf1.FormatType = ConditionalFormatType.CellValue;
            cf1.FirstFormula = "300";
            cf1.SecondFormula = "500";
            cf1.Operator = ComparisonOperatorType.Between;
            cf1.BackColor = Color.Yellow;
        }

        private void AddComparisonRule4(Worksheet sheet)
        {
            //Create conditional formatting rule
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range["A4:D4"]);
            IConditionalFormat cf1 = xcfs1.AddCondition();
            cf1.FormatType = ConditionalFormatType.CellValue;
            cf1.FirstFormula = "100";
            cf1.SecondFormula = "200";
            cf1.Operator = ComparisonOperatorType.NotBetween;
            //Set fill pattern type
            cf1.FillPattern = ExcelPatternType.ReverseDiagonalStripe;
            //Set foreground color
            cf1.Color = Color.FromArgb(255, 255, 0);

            //Set background color
            cf1.BackColor = Color.FromArgb(0, 255, 255);
        }

        private void FileViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
