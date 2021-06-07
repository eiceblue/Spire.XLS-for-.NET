using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

namespace GroupRowsAndColumns
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GroupRowsAndColumns.xls");
            Worksheet sheet = workbook.Worksheets[0];

            //Grouping rows
            sheet.GroupByRows(1,5,false);
            //Grouping columns
            sheet.GroupByColumns(1,3,false);

            workbook.SaveToFile("GroupRowsAndColumns.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer("GroupRowsAndColumns.xlsx");
		}

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }
	}
}
