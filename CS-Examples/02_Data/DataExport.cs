using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace DataExport
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

            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataExport.xlsx");
			//Initailize worksheet
			Worksheet sheet = workbook.Worksheets[0];

			this.dataGrid1.DataSource =  sheet.ExportDataTable();
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }



	}
}
