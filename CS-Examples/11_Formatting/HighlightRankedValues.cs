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

namespace HighlightRankedValues
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_6.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Apply conditional formatting to range ¡°D2:D10¡± to highlight the top 2 values.
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["D2:D10"]);
            IConditionalFormat format1 = xcfs.AddTopBottomCondition(TopBottomType.Top, 2);
            format1.FormatType = ConditionalFormatType.TopBottom;
            format1.BackColor = Color.Red;

            //Apply conditional formatting to range ¡°E2:E10¡± to highlight the bottom 2 values.
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range["E2:E10"]);
            IConditionalFormat format2 = xcfs1.AddTopBottomCondition(TopBottomType.Bottom,2);
            format2.FormatType = ConditionalFormatType.TopBottom;
            format2.BackColor = Color.ForestGreen;

            // Specify the output file name.
            String result = "Result-HighlightTopAndBottomRankedValues.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);
            // Dispose of the workbook object to release resources
            workbook.Dispose();

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
