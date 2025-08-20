using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace ReplaceTextInTextBox
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ReplaceTextInTextBox.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            string tag = "TAG_1$TAG_2";
            string replace = "Spire.XLS for .NET$Spire.XLS for JAVA";

            for (int i = 0; i < tag.Split('$').Length; i++)
            {
                // Replace text in textbox
                ReplaceTextInTextBox(sheet, "<" + tag.Split('$')[i] + ">", replace.Split('$')[i]);
            }

            // Specify the output file path
            string output = "ReplaceTextInTextBox_out.xlsx";

            // Save the modified workbook to a file
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
        }
        private void ReplaceTextInTextBox(Worksheet sheet, string sFind, string sReplace)
        {
            for (int i = 0; i < sheet.TextBoxes.Count; i++)
            {
                ITextBox tb = sheet.TextBoxes[i];
                if (!String.IsNullOrEmpty(tb.Text))
                {
                    if (tb.Text.Contains(sFind))
                    {
                        tb.Text = tb.Text.Replace(sFind, sReplace);
                    }
                }
            }
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
