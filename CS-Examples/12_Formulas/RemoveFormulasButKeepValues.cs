using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System.Runtime.Remoting.Lifetime;

namespace RemoveFormulasButKeepValues
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\RemoveFormulasButKeepValues.xlsx");

            //Loop through worksheets.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                //Loop through cells.
                foreach (CellRange cell in sheet.Range)
                {
                    //If the cell contain formula, get the formula value, clear cell content, and then fill the formula value into the cell.
                    if (cell.HasFormula)
                    {
                        Object value = cell.FormulaValue;
                        cell.Clear(ExcelClearOptions.ClearContent);
                        cell.Value2 = value;
                    }
                }
            }

            // Specify the name for the resulting Excel file
            String result = "Result-RemoveFormulasButKeepValues.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
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
