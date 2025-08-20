using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetHeaderFooter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook     
            Workbook workbook = new Workbook();

            // Load a Workbook from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetHeaderFooter.xlsx");

            // Get the first worksheet
            Worksheet Worksheet = workbook.Worksheets[0];


            // Set left header,"Arial Unicode MS" is font name, "18" is font size.
            Worksheet.PageSetup.LeftHeader = "&\"Arial Unicode MS\"&14 Spire.XLS for .NET ";

            // Set center footer 
            Worksheet.PageSetup.CenterFooter = "Footer Text";

            // Set view mode as  page layout view
            Worksheet.ViewMode = ViewMode.Layout;

            // Specify the file name for the resulting Excel file
            String result = "SetHeaderFooter_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
        }

        private void ExcelDocViewer(string fileName)
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
