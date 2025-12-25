using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Windows.Forms;

namespace AddShapeHyperlink
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

            workbook.LoadFromFile("..\\..\\..\\..\\..\\..\\Data\\AddShapeHyperlink.xlsx");

            // Get the reference to the first sheet in the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Get all the shapes in the sheet
            PrstGeomShapeCollection prstGeomShapeType = sheet.PrstGeomShapes;

            // Set the hyperlink for each shape
            for (int i = 0; i < prstGeomShapeType.Count; i++)
            {
                // Get the shape
                XlsPrstGeomShape shape = (XlsPrstGeomShape)prstGeomShapeType[i];

                // Set the hyperlink address
                shape.HyLink.Address = "https://www.e-iceblue.com/Download/download-excel-for-net-now.html";
            }

            // Specify the filename for the resulting Excel file
            String result = "AddShapeHyperlink-out.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object
            workbook.Dispose();

            // View the document using a file viewer
            FileViewer(result);
        }

        private void FileViewer(string fileName)
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
