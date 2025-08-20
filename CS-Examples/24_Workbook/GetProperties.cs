using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Collections;
using Spire.Xls.Core;

namespace GetProperties
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

            // Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\WorksheetSample1.xlsx");

            // Get the general excel properties
            BuiltInDocumentProperties properties1 = workbook.DocumentProperties;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Excel Properties:");
            for (int i = 0; i < properties1.Count; i++)
            {
                // Get property name
                string name = properties1[i].Name;
                // Get property vaule
                string value = properties1[i].Value.ToString();
                sb.AppendLine(name + ": " + value);
            }
            sb.AppendLine();

            //Get the custom properties
            ICustomDocumentProperties properties2 = workbook.CustomDocumentProperties;
            sb.AppendLine("Custom Properties:");
            for (int i = 0; i < properties2.Count; i++)
            {
                // Get property name
                string name = properties2[i].Name;
                // Get property vaule
                string value = properties2[i].Value.ToString();
                sb.AppendLine(name + ": " + value);
            }

            //Save the document
            string output = "GetProperties.txt";
            File.WriteAllText(output, sb.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the Excel file
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
