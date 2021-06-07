using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;

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

            String result = "Result-RemoveFormulasButKeepValues.xlsx";

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
