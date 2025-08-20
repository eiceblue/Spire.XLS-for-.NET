using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Spire.Xls.Core;

namespace RemoveCustomProperties
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

            // Load a excel document
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\templateAz.xlsx");

            // Retrieve a list of all custom document properties of the Excel file
            ICustomDocumentProperties customDocumentProperties = workbook.CustomDocumentProperties;

            // Remove "Editor" custom document property
            customDocumentProperties.Remove("Editor");

            // Save to file
            String result = "RemoveCustomProperties_result.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources 
            workbook.Dispose();

            // Launch the document
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
