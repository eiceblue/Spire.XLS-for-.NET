using System;
using System.Windows.Forms;
using Spire.Xls;
using Spire.Xls.Core;
using System.IO;
using Spire.Xls.Core.Spreadsheet.Collections;

namespace GetCustomPropertiesOfSheet
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

         private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load a Workbook from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GetCustomPropertiesOfSheet.xlsx");

            // Get the first sheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the custom properties of the first sheet
            ICustomPropertiesCollection customProperties = worksheet.CustomProperties;
            String information = "";
            for (int i = 0; i < customProperties.Count; i++)
            {
                XlsCustomProperty xcp = customProperties[i];
                string name = xcp.Name;
                information += "Name:" + name + "\n";
                string value = xcp.Value;
                information += "Value:" + value + "\n";
            }

            // Save the information to a .txt file
            String result = "GetCustomPropertiesOfSheet-out.txt";
            File.WriteAllText(result, information);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
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
