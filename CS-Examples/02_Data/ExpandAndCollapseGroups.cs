using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

namespace ExpandAndCollapseGroups
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Template_Xls_3.xlsx");

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Expand the grouped rows with ExpandCollapseFlags set to expand parent
            sheet.Range["A16:G19"].ExpandGroup(GroupByType.ByRows, ExpandCollapseFlags.ExpandParent);

            //Collapse the grouped rows
            sheet.Range["A10:G12"].CollapseGroup(GroupByType.ByRows);

            // Specify the name for the resulting Excel file
            String result = "Result-ExpandAndCollapseGroups.xlsx";

            //Save to file.
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object
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
