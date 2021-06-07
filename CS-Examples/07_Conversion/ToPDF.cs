using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ToPDF
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            using (Workbook workbook = new Workbook())
            {
                workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToPDF.xlsx");
                workbook.ConverterSetting.SheetFitToPage = true;
                workbook.SaveToFile("sample.pdf", FileFormat.PDF);
            }
            ExcelDocViewer("sample.pdf");
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
