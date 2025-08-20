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

namespace SetRowColorByConditionalFormat
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_4.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Select the range that you want to format.
            CellRange dataRange = sheet.AllocatedRange;

            //Set conditional formatting.
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(dataRange);
            IConditionalFormat format1 = xcfs.AddCondition();
            //Determines the cells to format.
            format1.FirstFormula = "=MOD(ROW(),2)=0";
            //Set conditional formatting type
            format1.FormatType = ConditionalFormatType.Formula;
            //Set the color.
            format1.BackColor = Color.LightSeaGreen;

            //Set the backcolor of the odd rows as Yellow.
            XlsConditionalFormats xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(dataRange);
            IConditionalFormat format2 = xcfs.AddCondition(); 
            format2.FirstFormula = "=MOD(ROW(),2)=1";
            format2.FormatType = ConditionalFormatType.Formula;
            format2.BackColor = Color.Yellow;

            // Specify the name for the resulting Excel file
            String result = "Result-SetRowColorWithConditionalFormatting.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
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
