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

namespace CopyDataWithStyle
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
      
            //Get the default first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            //Set the values for some cells.
            CellRange cells = worksheet.Range["A1:J50"];
            for (int i = 1; i <= 10; i++)
            {
                for (int j = 1; j <= 8; j++)
                {
                    string text = string.Format((i - 1).ToString() + "," + (j - 1).ToString());
                    cells[i, j].Text = text;
                }
            }
            //Get a source range (A1:D3).
            CellRange srcRange = worksheet.Range["A1:D3"];

            //Create a style object.
            CellStyle style = workbook.Styles.Add("style");

            //Specify the font attribute.
            style.Font.FontName = "Calibri";

            //Specify the shading color.
            style.Font.Color = Color.Red;

            //Specify the border attributes.
            style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            style.Borders[BordersLineType.EdgeTop].Color = Color.Blue;
            style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            style.Borders[BordersLineType.EdgeBottom].Color = Color.Blue;
            style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            style.Borders[BordersLineType.EdgeTop].Color = Color.Blue;
            style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            style.Borders[BordersLineType.EdgeRight].Color = Color.Blue;
            srcRange.CellStyleName = style.Name;

            //Set the destination range
            CellRange destRange = worksheet.Range["A12:D14"];

            //Copy the range data with style
            srcRange.Copy(destRange, true, true);
     
            //String for output file 
            String outputFile = "Output.xlsx";

            //Save the file
            workbook.SaveToFile(outputFile, ExcelVersion.Version2013);

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
