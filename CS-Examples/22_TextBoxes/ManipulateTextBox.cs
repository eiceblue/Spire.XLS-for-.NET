using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace ManipulateTextBox
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ManipulateTextBoxControl.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first textbox from the worksheet
            ITextBox tb = sheet.TextBoxes[0];

            // Change the text of the textbox
            tb.Text = "Spire.XLS for .NET";

            // Set the alignment of the textbox as center
            tb.HAlignment = CommentHAlignType.Center;
            tb.VAlignment = CommentVAlignType.Center;

            // Specify the output file path
            string output = "ManipulateTextBoxControl_out.xlsx";
            // Save the modified workbook to a file
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
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
