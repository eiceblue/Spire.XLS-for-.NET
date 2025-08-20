using System;
using System.Drawing;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;

namespace AddOvalShape
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

            //Add oval shape1
            IOvalShape ovalShape1 = sheet.OvalShapes.AddOval(11, 2, 100, 100);
            ovalShape1.Line.Weight = 0;
            //Fill shape with solid color
            ovalShape1.Fill.FillType = ShapeFillType.SolidColor;
            ovalShape1.Fill.ForeColor = Color.DarkCyan;

            //Add oval shape2
            IOvalShape ovalShape2 = sheet.OvalShapes.AddOval(11, 5, 100, 100);
            ovalShape2.Line.Weight = 1;
            //Fill shape with picture
            ovalShape2.Line.DashStyle = ShapeDashLineStyleType.Solid;
            ovalShape2.Fill.CustomPicture(@"..\..\..\..\..\..\Data\logo.png");

            //Save the document
            string output = "AddOvalShape_out.xlsx";
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
