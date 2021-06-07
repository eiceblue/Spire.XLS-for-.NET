using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet.Shapes;

namespace SetImageOffsetOfChart
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, System.EventArgs e)
		{
            //Create a workbook.
			Workbook workbook = new Workbook();

            //Load the file from disk.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_1.xlsx");

            //Get the first worksheet.
			      Worksheet sheet = workbook.Worksheets[0];
            Worksheet sheet1 = workbook.Worksheets.Add("Contrast");

            //Add chart1 and background image to sheet1 as comparision.
            Chart chart1 = sheet1.Charts.Add(ExcelChartType.ColumnClustered);
            chart1.DataRange = sheet.Range["D1:E8"];
            chart1.SeriesDataFromRange = false;

            //Chart Position.
            chart1.LeftColumn = 1;
            chart1.TopRow = 11;
            chart1.RightColumn = 8;
            chart1.BottomRow = 33;

            //Add picture as background.
            chart1.ChartArea.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\Background.png"), "None");
            
            chart1.ChartArea.Fill.Tile = false;
        
            //Set the image offset.  
            chart1.ChartArea.Fill.PicStretch.Left = 20;
            chart1.ChartArea.Fill.PicStretch.Top = 20;
            chart1.ChartArea.Fill.PicStretch.Right = 5;
            chart1.ChartArea.Fill.PicStretch.Bottom = 5;

            String result = "Result-SetImageOffsetOfChart.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            //Launch the MS Excel file.
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
