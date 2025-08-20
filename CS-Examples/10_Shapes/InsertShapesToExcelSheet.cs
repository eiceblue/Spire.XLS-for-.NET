using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace InsertShapesToExcelSheet
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

            //Get the first worksheet.
			Worksheet sheet = workbook.Worksheets[0];

            //Add a triangle shape.
            IPrstGeomShape triangle = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, PrstGeomShapeType.Triangle);
            //Fill the triangle with solid color.
            triangle.Fill.ForeColor = Color.Yellow;
            triangle.Fill.FillType = ShapeFillType.SolidColor;

            //Add a heart shape.
            IPrstGeomShape heart = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, PrstGeomShapeType.Heart);          
            //Fill the heart with gradient color.
            heart.Fill.ForeColor = Color.Red;
            heart.Fill.FillType = ShapeFillType.Gradient;

            //Add an arrow shape with default color.
            IPrstGeomShape arrow = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, PrstGeomShapeType.CurvedRightArrow);

            //Add a cloud shape.
            IPrstGeomShape cloud = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, PrstGeomShapeType.Cloud);
            //Fill the cloud with custom picture
            cloud.Fill.CustomPicture(Image.FromFile(@"..\..\..\..\..\..\Data\SpireXls.png"), "SpireXls.png");
            cloud.Fill.FillType = ShapeFillType.Picture;

            //Save to file.
            String result = "Result-InsertShapesToExcelSheet.xlsx";
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
