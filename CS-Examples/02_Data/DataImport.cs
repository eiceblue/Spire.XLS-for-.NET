using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace DataImport
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

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Import date to data table
            sheet.InsertDataTable((DataTable)this.dataGrid1.DataSource,true,1,1,-1,-1);

			// Set body style
			CellStyle oddStyle = workbook.Styles.Add("oddStyle");
			oddStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
			oddStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
			oddStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
			oddStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
			oddStyle.KnownColor = ExcelColors.LightGreen1;

			CellStyle evenStyle = workbook.Styles.Add("evenStyle");
			evenStyle.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
			evenStyle.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
			evenStyle.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
			evenStyle.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
			evenStyle.KnownColor = ExcelColors.LightTurquoise;

			foreach( CellRange range in  sheet.AllocatedRange.Rows)
			{
				if (range.Row % 2 == 0)
					range.CellStyleName = evenStyle.Name;
			    else
					range.CellStyleName = oddStyle.Name;
			}

			// Set header style
			CellStyle styleHeader = sheet.Rows[0].Style;
			styleHeader.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
			styleHeader.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
			styleHeader.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
			styleHeader.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
			styleHeader.VerticalAlignment = VerticalAlignType.Center;
			styleHeader.KnownColor = ExcelColors.Green;
			styleHeader.Font.KnownColor = ExcelColors.White;
			styleHeader.Font.IsBold = true;

            // Auto-fit the columns to adjust their widths
            sheet.AllocatedRange.AutoFitColumns();
            // Auto-fit the rows to adjust their height
            sheet.AllocatedRange.AutoFitRows();

			// Set row heihgt
            sheet.Rows[0].RowHeight = 20;

            // Specify the name for the resulting Excel file
            string result = "DataImport_out.xls";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result);

			// Launch the file
            ExcelDocViewer(result);
		}
		private void Form1_Load(object sender, System.EventArgs e)
		{
            // Create a new workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\DataImport.xls");
			//Initailize worksheet
			Worksheet sheet = workbook.Worksheets[0];

            //Export the first sheet data to dataTable 
            this.dataGrid1.DataSource =  sheet.ExportDataTable();
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
