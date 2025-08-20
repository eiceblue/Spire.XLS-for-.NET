using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DecryptWorkbook
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            string fileName = @"..\..\..\..\..\..\Data\DecryptWorkbook.xlsx";

            //Detect if the Excel workbook is password protected.
            bool value = Workbook.IsPasswordProtected(fileName);

            if (value)
            {
                //Load a file with the password specified
                Workbook workbook = new Workbook();
                workbook.OpenPassword = "eiceblue";
                workbook.LoadFromFile(fileName);

                //Decrypt workbook
                workbook.UnProtect();

                //Save the document
                workbook.SaveToFile("DecryptWorkbook_result.xlsx", ExcelVersion.Version2010);

                // Dispose of the workbook object to release resources
                workbook.Dispose();
            }



            ExcelDocViewer("DecryptWorkbook_result.xlsx");
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
