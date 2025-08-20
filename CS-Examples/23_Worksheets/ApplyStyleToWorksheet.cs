using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;

namespace ApplyStyleToWorksheet
{

	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }

		private void btnRun_Click(object sender, EventArgs e)
		{
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Load the Excel document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample1.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Create a cell style
            CellStyle style = workbook.Styles.Add("newStyle");
            style.Color = Color.LightBlue;
            style.Font.Color = Color.White;
            style.Font.Size = 15;
            style.Font.IsBold = true;

            // Apply the style to the first worksheet
            sheet.ApplyStyle(style);

            // Save the modified workbook to a file
            string output = "ApplyStyleToWorksheet.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
            ExcelDocViewer(output);
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
