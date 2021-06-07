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

namespace SetTrafficLightsIcons
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

            //Add a worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            //Add some data to the Excel sheet cell range and set the format for them.
            sheet.Range["A1"].Text = "Traffic Lights";
            sheet.Range["A2"].NumberValue = 0.95;
            sheet.Range["A2"].NumberFormat = "0%";
            sheet.Range["A3"].NumberValue = 0.5;
            sheet.Range["A3"].NumberFormat = "0%";
            sheet.Range["A4"].NumberValue = 0.1;
            sheet.Range["A4"].NumberFormat = "0%";
            sheet.Range["A5"].NumberValue = 0.9;
            sheet.Range["A5"].NumberFormat = "0%";
            sheet.Range["A6"].NumberValue = 0.7;
            sheet.Range["A6"].NumberFormat = "0%";
            sheet.Range["A7"].NumberValue = 0.6;
            sheet.Range["A7"].NumberFormat = "0%";

            //Set the height of row and width of column for Excel cell range.
            sheet.AllocatedRange.RowHeight = 20;
            sheet.AllocatedRange.ColumnWidth = 25;

            //Add a conditional formatting.
            XlsConditionalFormats conditional = sheet.ConditionalFormats.Add();
            conditional.AddRange(sheet.AllocatedRange);
            IConditionalFormat format1 = conditional.AddCondition();

            //Add a conditional formatting of cell range and set its type to CellValue.
            format1.FormatType = ConditionalFormatType.CellValue;
            format1.FirstFormula = "300";
            format1.Operator = ComparisonOperatorType.Less;
            format1.FontColor = Color.Black;
            format1.BackColor = Color.LightSkyBlue;

            //Add a conditional formatting of cell range and set its type to IconSet.
            conditional.AddRange(sheet.AllocatedRange);
            IConditionalFormat format = conditional.AddCondition();
            format.FormatType = ConditionalFormatType.IconSet;
            format.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;

            String result = "Result-SetTrafficLightsIcons.xlsx";

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
