using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace LinkToContentProperty
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AccessDocumentProperties.xlsx");

            // Add a custom document property
            workbook.CustomDocumentProperties.Add("Test", "MyNamedRange");

            // Get the added document property
            ICustomDocumentProperties properties = workbook.CustomDocumentProperties;
            DocumentProperty property = (DocumentProperty)properties["Test"];

            // Link to content 
            property.LinkToContent = true;

            // Save the document
            string result = "LinkToContentProperty_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
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
