using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Drawing.Printing;

namespace PrintExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PrintExcel.xlsx");

            // Access the printer settings of the workbook's print document
            PrinterSettings settings = workbook.PrintDocument.PrinterSettings;

            // Specify the range of pages to be printed (from page 0 to page 1)
            settings.FromPage = 0;
            settings.ToPage = 1;

            // Use the default printer to print
            workbook.PrintDocument.Print();

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
