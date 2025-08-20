using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace AddLineShape
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Add shape line1
            ILineShape line1 = sheet.Lines.AddLine(10, 2, 200, 1, LineShapeType.Line);
            //Set dash style type
            line1.DashStyle = ShapeDashLineStyleType.Solid;
            //Set color
            line1.Color = Color.CadetBlue;
            //Set weight
            line1.Weight = 2f;
            //Set end arrow style type
            line1.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

            //Add shape line2
            ILineShape line2 = sheet.Lines.AddLine(12, 2, 200, 1, LineShapeType.CurveLine);
            line2.DashStyle = ShapeDashLineStyleType.Dotted;
            line2.Color = Color.OrangeRed;
            line2.Weight = 2f;

            //Add shape line3
            ILineShape line3 = sheet.Lines.AddLine(14, 2, 200, 1, LineShapeType.ElbowLine);
            line3.DashStyle = ShapeDashLineStyleType.DashDotDot;
            line3.Color = Color.Purple;
            line3.Weight = 2f;

            //Add shape line4
            ILineShape line4 = sheet.Lines.AddLine(16, 2, 200, 1, LineShapeType.LineInv);
            line4.DashStyle = ShapeDashLineStyleType.Dashed;
            line4.Color = Color.Green;
            line4.Weight = 2f;

            //Save the document
            string output = "InsertLineShape_out.xlsx";
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
