using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace TextBoxWithWrapText
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

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\TextBoxSampleB.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Get the text box
            XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

            // Set wrap text
            shape.IsWrapText = true;

            // Specify the output filename for the workbook
            string output = "TextBoxWithWrapText.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // View the document
            FileViewer(output);
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
