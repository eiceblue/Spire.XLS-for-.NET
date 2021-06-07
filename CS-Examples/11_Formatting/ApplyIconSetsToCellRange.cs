using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core;

namespace ApplyIconSetsToCellRange
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

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Insert data to cell range from A1 to C4.
            sheet.Range["A1"].NumberValue = 582;
            sheet.Range["A2"].NumberValue = 234;
            sheet.Range["A3"].NumberValue = 314;
            sheet.Range["A4"].NumberValue = 50;
            sheet.Range["B1"].NumberValue = 150;
            sheet.Range["B2"].NumberValue = 894;
            sheet.Range["B3"].NumberValue = 560;
            sheet.Range["B4"].NumberValue = 900;
            sheet.Range["C1"].NumberValue = 134;
            sheet.Range["C2"].NumberValue = 700;
            sheet.Range["C3"].NumberValue = 920;
            sheet.Range["C4"].NumberValue = 450;
            sheet.AllocatedRange.RowHeight = 15;
            sheet.AllocatedRange.ColumnWidth = 17;

            //Add icon sets.
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.AllocatedRange);
            IConditionalFormat format = xcfs.AddCondition(); 
            format.FormatType = ConditionalFormatType.IconSet;
            format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;

            String result = "Result-ApplyIconSetsToDataRange.xlsx";

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
