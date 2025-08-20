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
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load an existing file from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorkbookToHTML.xlsx");

            // Get the first worksheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Find the string "Address" in the worksheet
            CellRange cr = sheet.FindString("Address", false, false);

            // Delete the row that includes the found string
            sheet.DeleteRow(cr.Row);

            // Save the modified workbook to a new file
            workbook.SaveToFile("RemoveRowBasedOnKeyword.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
