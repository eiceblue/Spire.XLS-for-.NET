using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace SetHeightAndWidth
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SetHeightAndWidth.xls");
              
            Worksheet worksheet = workbook.Worksheets[0];
            // Setting the width to 30
            worksheet.SetColumnWidth(4, 30);
            // Setting the height to 30
            worksheet.SetRowHeight(4,30);

            string result="SetHeightAndWidth_out.xlsx";
            workbook.SaveToFile(result,ExcelVersion.Version2010);
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
