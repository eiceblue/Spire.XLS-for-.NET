using System;
using System.Windows.Forms;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core;
using System.Text;
using System.IO;

namespace GetShapeLinkedCellRange
{
	public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
		{
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Create a StringBuilder to store the cell addresses.
            StringBuilder sb = new StringBuilder();

            // Load an existing Excel file.
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CellLinkedRangeLocal.xlsx");

            // Get the first worksheet from the workbook.
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of preset geometric shapes in the sheet.
            PrstGeomShapeCollection prstGeomShapeCollection = sheet.PrstGeomShapes;

            // Get a specific shape by its name.
            IPrstGeomShape shape = prstGeomShapeCollection["Yesterday"];

            // Get the range address of the cell linked to the shape.
            string cellAddress = shape.LinkedCell.RangeAddress;

            // Append the cell address to the StringBuilder.
            sb.Append(cellAddress + "\n");

            // Get another shape by its name.
            shape = prstGeomShapeCollection["NewShapes"];

            // Get the range address of the cell linked to the shape.
            cellAddress = shape.LinkedCell.RangeAddress;

            // Append the cell address to the StringBuilder.
            sb.Append(cellAddress);

            // Write the content of the StringBuilder to an output text file.
            File.WriteAllText("output.txt", sb.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("output.txt");
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
