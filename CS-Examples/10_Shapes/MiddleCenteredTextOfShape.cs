using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace MiddleCenteredTextOfShape
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

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            IPrstGeomShape rect = sheet.PrstGeomShapes.AddPrstGeomShape(8, 2, 300, 300, PrstGeomShapeType.Rect);

            // Set the fill color of the rectangle to white (solid color)
            rect.Fill.ForeColor = Color.White;
            rect.Fill.FillType = ShapeFillType.SolidColor;

            // Set the text content of the rectangle
            rect.Text = "E-iceblue";

            // Set the vertical alignment of the text to middle-centered
            rect.TextVerticalAlignment = ExcelVerticalAlignment.MiddleCentered;

            // Save the workbook to a file
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
            FileViewer("result.xlsx");
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
