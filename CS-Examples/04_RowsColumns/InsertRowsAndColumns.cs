using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace InsertRowsAndColumns
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\InsertRowsAndColumns.xls");
              
            Worksheet worksheet = workbook.Worksheets[0];
            //Inserting a row into the worksheet 
            worksheet.InsertRow(2);
            //Inserting a column into the worksheet 
            worksheet.InsertColumn(2);
            //Inserting multiple rows into the worksheet
            worksheet.InsertRow(5, 2);
            //Inserting multiple columns into the worksheet
            worksheet.InsertColumn(5, 2);

            string result="InsertRowsAndColumns_out.xlsx";
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
