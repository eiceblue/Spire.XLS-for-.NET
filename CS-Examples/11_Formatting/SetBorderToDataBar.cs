using Spire.Xls;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet.Collections;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace SetBorderToDataBar
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_9.xlsx");

            // Get the first sheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Get the data bar format from the first conditional format
            XlsConditionalFormats xcfs = sheet.ConditionalFormats[0];
            IConditionalFormat cf = xcfs[0];
            Spire.Xls.DataBar dataBar1 = cf.DataBar;

            // Set the border type and color for the data bar format
            dataBar1.BarBorder.Type = Spire.Xls.Core.Spreadsheet.ConditionalFormatting.DataBarBorderType.DataBarBorderSolid;
            dataBar1.BarBorder.Color = Color.Red;

            // Set a new data bar format to cell E1
            sheet["E1"].NumberValue = 200;
            XlsConditionalFormats xcfs2 = sheet.ConditionalFormats.Add();
            xcfs2.AddRange(sheet.Range["E1"]);
            IConditionalFormat cf2 = xcfs2.AddCondition();
            cf2.FormatType = ConditionalFormatType.DataBar;
            cf2.DataBar.BarBorder.Type = Spire.Xls.Core.Spreadsheet.ConditionalFormatting.DataBarBorderType.DataBarBorderSolid;
            cf2.DataBar.BarBorder.Color = Color.Red;
            cf2.DataBar.BarColor = Color.GreenYellow;

            // Save the modified workbook to a file
            String result = "SetBorderToDataBar_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer(result);
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
