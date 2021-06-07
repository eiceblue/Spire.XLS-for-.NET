using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Text;
using System.IO;

namespace HtmlToExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //File path
            string filePath = @"..\..\..\..\..\..\Data\HtmlToExcel.html";

            //Create a workbook
            Workbook workbook = new Workbook();

            //Load html
            workbook.LoadFromHtml(filePath);

            //Save to Excel file
            string result = "HtmlToExcel_result.xlsx";

            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //Launch the file
            ExcelDocViewer(result);
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
