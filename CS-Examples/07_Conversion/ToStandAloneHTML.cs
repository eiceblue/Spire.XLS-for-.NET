using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ToStandAloneHTML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Load an existing Excel document from the specified path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToCSV.xlsx");

            // Set the HTMLOptions to create a standalone HTML file.
            HTMLOptions.Default.IsStandAloneHtmlFile=true;

            // Specify the filename for the output HTML file.
            String outputFile = "Output.html";

            // Create a FileStream object to write the HTML file.
            FileStream fileStream = new FileStream(outputFile, FileMode.Create);

            // Save the Excel document as an HTML file to the FileStream.
            workbook.SaveToStream(fileStream, FileFormat.HTML);

            // Close the FileStream.
            fileStream.Close();

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //view the document
            ExcelDocViewer(outputFile);
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
