using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveRowBasedOnKeyword
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorkbookToHTML.xlsx");
          
            Worksheet sheet = workbook.Worksheets[0];

            //Find the string
            CellRange cr = sheet.FindString("Address", false, false);

            //Delete the row which includes the string
            sheet.DeleteRow(cr.Row);

            //Save to file
            workbook.SaveToFile("RemoveRowBasedOnKeyword.xlsx", ExcelVersion.Version2010);

            //View the document
            FileViewer("RemoveRowBasedOnKeyword.xlsx");
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
