using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;
using System.Text;
using System.IO;

namespace AccessDocumentProperties
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

            // Create string builder
            StringBuilder builder = new StringBuilder();

            // Get all document properties
            ICustomDocumentProperties properties = workbook.CustomDocumentProperties;

            // Access document property by property name
            DocumentProperty property1 = (DocumentProperty)properties["Editor"];
            builder.AppendLine(property1.Name + " " + property1.Value);

            // Access document property by property index
            DocumentProperty property2 = (DocumentProperty)properties[0];
            builder.AppendLine(property2.Name + " " + property2.Value);

            // Save to txt file
            string result = "AccessDocumentProperties_out.txt";
            File.WriteAllText(result, builder.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the file 
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
