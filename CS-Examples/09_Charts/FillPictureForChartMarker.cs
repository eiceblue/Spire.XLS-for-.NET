using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace FillPictureForChartMarker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {

            // Specify the path of the input Excel file.
            string inputFile = @"..\..\..\..\..\..\Data\FillChartMarker.xlsx";

            // Specify the path of the image file.
            string imageFile = @"..\..\..\..\..\..\Data\E-iceblueLogo.png";

            // Create a new Workbook object.
            Workbook workbook = new Workbook();

            // Load the input Excel file.
            workbook.LoadFromFile(inputFile);

            // Get the first worksheet from the workbook.
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the first chart from the worksheet.
            Chart chart = worksheet.Charts[0];

            // Set the line color of series 1 to yellow.
            chart.Series[0].Format.LineProperties.Color = Color.Yellow;

            // Set the marker style of series 1 to picture.
            chart.Series[0].Format.MarkerStyle = ChartMarkerType.Picture;

            // Get the marker fill for series 1.
            IShapeFill markerFill1 = chart.Series[0].DataFormat.MarkerFill;

            // Set the custom picture for the marker fill of series 1.
            markerFill1.CustomPicture(imageFile);

            // Get the marker fill for series 2.
            IShapeFill markerFill2 = chart.Series[1].DataFormat.MarkerFill;

            // Set the line color of series 2 to red.
            chart.Series[1].Format.LineProperties.Color = Color.Red;

            // Set the texture of the marker fill for series 2 to granite.
            markerFill2.Texture = GradientTextureType.Granite;

            // Set the line color of series 1 to blue.
            chart.Series[0].Format.LineProperties.Color = Color.Blue;

            // Get the marker fill for series 3.
            IShapeFill markerFill3 = chart.Series[2].DataFormat.MarkerFill;

            // Set the pattern of the marker fill for series 3 to 10% gradient
            markerFill3.Pattern = GradientPatternType.Pat10Percent;

            // Set the foreground color of the marker fill for series 3 to light gray.
            markerFill3.ForeColor = Color.LightGray;

            // Set the background color of the marker fill for series 3 to orange.
            markerFill3.BackColor = Color.Orange;

            // Save the modified workbook to a new Excel file.
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
