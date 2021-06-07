using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using Spire.Xls.Core;

namespace SetShadowStyleForShape
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

            //Add an ellipse shape.
            IPrstGeomShape ellipse = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, PrstGeomShapeType.Ellipse);

            //Set the shadow style for the ellipse.
            ellipse.Shadow.Angle = 90;
            ellipse.Shadow.Distance = 10;
            ellipse.Shadow.Size = 150;
            ellipse.Shadow.Color = Color.Gray;
            ellipse.Shadow.Blur = 30;
            ellipse.Shadow.Transparency = 1;
            ellipse.Shadow.HasCustomStyle = true;

            String result = "Result-SetShapeShadowStyleForNewFile.xlsx";

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
