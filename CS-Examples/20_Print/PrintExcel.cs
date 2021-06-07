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
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\PrintExcel.xlsx");
            PrinterSettings settings = workbook.PrintDocument.PrinterSettings;
            settings.FromPage = 0;
            settings.ToPage = 1;
            //Use the default printer to print
            workbook.PrintDocument.Print();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
