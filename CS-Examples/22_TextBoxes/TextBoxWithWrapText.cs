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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk          
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\TextBoxSampleB.xlsx");
       
            Worksheet sheet = workbook.Worksheets[0];
            //Get the text box
            XlsTextBoxShape shape = sheet.TextBoxes[0] as XlsTextBoxShape;

            //Set wrap text
            shape.IsWrapText = true;

            //Save the document
            string output = "TextBoxWithWrapText.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            //View the document
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
