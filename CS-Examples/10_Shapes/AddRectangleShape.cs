using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace AddRectangleShape
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

            //Add rectangle shape 1------Rect
            IRectangleShape rect1=sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, RectangleShapeType.Rect);
            rect1.Line.Weight = 1;
            //Fill shape with solid color
            rect1.Fill.FillType = ShapeFillType.SolidColor;
            rect1.Fill.ForeColor = Color.DarkGreen;

            //Add rectangle shape 2------RoundRect
            IRectangleShape rect2 = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100, RectangleShapeType.RoundRect);
            rect2.Line.Weight = 1;
            rect2.Fill.FillType = ShapeFillType.SolidColor;
            rect2.Fill.ForeColor = Color.DarkCyan;

            //Save the document
            string output = "AddRectangleShape_out.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

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
