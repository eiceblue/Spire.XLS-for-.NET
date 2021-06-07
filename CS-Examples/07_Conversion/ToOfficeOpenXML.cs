using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

using Spire.Xls;

namespace ToOfficeOpenXML
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
			Worksheet sheet = workbook.Worksheets[0];
            sheet.Range["A1"].Text = "Hello World";
			sheet.Range["B1"].Style.KnownColor = ExcelColors.Gray25Percent;
			sheet.Range["C1"].Style.KnownColor= ExcelColors.Gold;
			workbook.SaveAsXml("sample.xml");

			System.Diagnostics.Process.Start(Path.Combine(Application.StartupPath,"Sample.xml"));
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
	}
}
