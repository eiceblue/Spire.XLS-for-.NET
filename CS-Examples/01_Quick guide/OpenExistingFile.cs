using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace OpenExistingFile
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

            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz2.xlsx");

            //Add a new sheet, named MySheet
            Worksheet sheet = workbook.Worksheets.Add("MySheet");

            //Get the reference of "A1" cell from the cells collection of a worksheet
            sheet.Range["A1"].Text = "Hello World";


            String result = "OpenExistingFile_result.xlsx";

            //Save to file
            workbook.SaveToFile(result, ExcelVersion.Version2010);

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
