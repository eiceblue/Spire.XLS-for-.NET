using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Spire.Xls;
using Spire.Xls.Charts;
using System.Text;
using System.Collections.Generic;

namespace VerifyDataByValidation
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Sample.xlsx");

            //Get first worksheet of the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            //Cell B4 has the Decimal Validation
            CellRange cell = worksheet.Range["B4"];

            //Get the valditation of this cell
            Validation validation = cell.DataValidation;

            //Get the specified data range
            double minimum = double.Parse(validation.Formula1);
            double maximum = double.Parse(validation.Formula2);

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set different numbers for the cell
            for (int i = 5; i < 100; i=i+40 )
            {
                cell.NumberValue = i;
                string result=null;
                //Verify 
                if (cell.NumberValue < minimum || cell.NumberValue > maximum)
                {
                    //Set string format for displaying
                    result = string.Format("Is input "+ i +" a valid value for this Cell: false");
                }
                else
                {
                    //Set string format for displaying
                    result = string.Format("Is input " + i + " a valid value for this Cell: true");
                }
                //Add result string to StringBuilder
                content.AppendLine(result);
            }
            //String for output file 
            String outputFile = "Output.txt";

            //Save them to a txt file
            File.WriteAllText(outputFile, content.ToString());

            //Launching the output file.
            Viewer(outputFile);
		}
		private void Viewer( string fileName )
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
