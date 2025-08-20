using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AddCustomProperties
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
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\AddCustomProperties.xlsx");

            // Load the document from disk
            workbook.CustomDocumentProperties.Add("_MarkAsFinal", true);
            
            // Add other custom properties to the workbook
            workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue");
            workbook.CustomDocumentProperties.Add("Phone number", 81705109);
            workbook.CustomDocumentProperties.Add("Revision number", 7.12);
            workbook.CustomDocumentProperties.Add("Revision date", DateTime.Now);

            // Save the document
            workbook.SaveToFile("AddCustomProperties_result.xlsx", FileFormat.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer("AddCustomProperties_result.xlsx");
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
