using Spire.Xls;
using Spire.Xls.Core.MergeSpreadsheet.Collections;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace GroupShapes
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

            // Get the first sheet from the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Add shapes to the worksheet
            IPrstGeomShape shape1 = worksheet.PrstGeomShapes.AddPrstGeomShape(1, 3, 50, 50, PrstGeomShapeType.RoundRect);
            IPrstGeomShape shape2 = worksheet.PrstGeomShapes.AddPrstGeomShape(5, 3, 50, 50, PrstGeomShapeType.Triangle);

            // Group the shapes
            GroupShapeCollection groupShapeCollection = worksheet.GroupShapeCollection;
            groupShapeCollection.Group(new Spire.Xls.Core.IShape[] { shape1, shape2 });

            // Save the workbook to a file
            string result = "GroupShape.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //View the document
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
