using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;


namespace AddArrowLineToExcelFile
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

            //Add a Double Arrow and fill the line with solid color.
            var line = sheet.TypedLines.AddLine();
            line.Top = 10;
            line.Left = 20;
            line.Width = 100;
            line.Height = 0;
            line.Color = Color.Blue;
            line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow;
            line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

            //Add an Arrow and fill the line with solid color.
            var line_1 = sheet.TypedLines.AddLine();
            line_1.Top = 50;
            line_1.Left = 30;
            line_1.Width = 100;
            line_1.Height = 100;
            line_1.Color = Color.Red;
            line_1.BeginArrowHeadStyle = ShapeArrowStyleType.LineNoArrow;
            line_1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

            //Add an Elbow Arrow Connector.
            Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape line3 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
            line3.LineShapeType = LineShapeType.ElbowLine;
            line3.Width = 30;
            line3.Height = 50;
            line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;
            line3.Top = 100;
            line3.Left = 50;

            //Add an Elbow Double-Arrow Connector.
            Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape line2 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
            line2.LineShapeType = LineShapeType.ElbowLine;
            line2.Width = 50;
            line2.Height = 50;
            line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;
            line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow;
            line2.Left = 120;
            line2.Top = 100;
            
            //Add a Curved Arrow Connector.
            line3 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
            line3.LineShapeType = LineShapeType.CurveLine;
            line3.Width = 30;
            line3.Height = 50;
            line3.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen;
            line3.Top = 100;
            line3.Left = 200;

            //Add a Curved Double-Arrow Connector.
            line2 = sheet.TypedLines.AddLine() as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
            line2.LineShapeType = LineShapeType.CurveLine;
            line2.Width = 30;
            line2.Height = 50;
            line2.EndArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen;
            line2.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowOpen;
            line2.Left = 250;
            line2.Top = 100;

            //Save to file.
            String result = "Result-AddArrowLineToExcelFile.xlsx";
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
