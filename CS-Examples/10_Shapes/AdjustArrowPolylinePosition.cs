using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.Shapes;
using System;
using System.Windows.Forms;

namespace AdjustArrowPolylinePosition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Draw an elbow arrow
            XlsLineShape line = worksheet.TypedLines.AddLine(5, 5, 100, 100, LineShapeType.ElbowLine) as Spire.Xls.Core.Spreadsheet.Shapes.XlsLineShape;
            line.EndArrowHeadStyle = ShapeArrowStyleType.LineNoArrow;
            line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrow;
            GeomertyAdjustValue ad = line.ShapeAdjustValues.AddAdjustValue(GeomertyAdjustValueFormulaType.LiteralValue);

            // When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
            ad.SetFormulaParameter(new double[] {-50});

            // Save to file
            String result = "AdjustArrowPolylinePosition.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
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
