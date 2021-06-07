using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace TillPicAsTextureInShape
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\TillPicAsTextureInShape.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Get the first shape
            IPrstGeomShape shape = sheet.PrstGeomShapes[0];

            //Fill shape with texture
            shape.Fill.FillType = ShapeFillType.Texture;

            //Custom texture with picture
            shape.Fill.CustomTexture(@"..\..\..\..\..\..\Data\logo.png");

            //Tile pciture as texture 
            shape.Fill.Tile = true;

            //Save the document
            string output = "TillPicAsTextureInShape_out.xlsx";
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
