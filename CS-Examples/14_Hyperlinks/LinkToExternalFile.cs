using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace LinkToExternalFile
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get cell "A1"
            CellRange range = sheet.Range[1, 1];

            // Add a hyperlink within the specified range
            HyperLink hyperlink = sheet.HyperLinks.Add(range);

            // Set the hyperlink type
            hyperlink.Type = HyperLinkType.File;

            // Set the display text for the hyperlink
            hyperlink.TextToDisplay = "Link To External File";

            // Set the file address for the hyperlink
            hyperlink.Address = "..\\..\\..\\..\\..\\..\\Data\\SampleB_4.xlsx";

            // Specify the file name for the resulting file
            string result = "result.xlsx";

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
