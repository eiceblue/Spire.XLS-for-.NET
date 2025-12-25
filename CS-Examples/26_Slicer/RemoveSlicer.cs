using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RemoveSlicer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Workbook instance
            Workbook wb = new Workbook();

            // Load an existing Excel file from the specified path
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\SlicerTemplate.xlsx");

            // Get the first worksheet in the workbook
            Worksheet worksheet = wb.Worksheets[0];

            // Get the slicer collection from the worksheet
            XlsSlicerCollection slicers = worksheet.Slicers;

            // Example: Remove the first slicer in the collection 
            // slicers.RemoveAt(0);

            // Clear all slicers from the collection
            slicers.Clear();

            // Save the modified workbook to a new file with Excel 2013 version format
            wb.SaveToFile("RemoveSlicer.xlsx", ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            wb.Dispose();

            // Launch the file
            ExcelDocViewer("RemoveSlicer.xlsx");
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
