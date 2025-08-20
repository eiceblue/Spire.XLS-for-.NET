using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SetDataDirection
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

            // Load an existing workbook from a file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner2.xlsx");

            // Create a DataTable named "data"
            DataTable dt = new DataTable("data");

            // Define a column named "value" in the DataTable
            dt.Columns.Add(new DataColumn("value", typeof(string)));

            // Create three new rows for the DataTable
            DataRow drName1 = dt.NewRow();
            DataRow drName2 = dt.NewRow();
            DataRow drName3 = dt.NewRow();

            // Set values for the "value" column in each row
            drName1["value"] = "Text1";
            drName2["value"] = "Text2";
            drName3["value"] = "Text3";

            // Add the rows to the DataTable
            dt.Rows.Add(drName1);
            dt.Rows.Add(drName2);
            dt.Rows.Add(drName3);

            // Add the DataTable to the Marker Designer with the parameter name "data"
            workbook.MarkerDesigner.AddDataTable("data", dt);

            // Apply the changes made in the Marker Designer
            workbook.MarkerDesigner.Apply();

            // Specify the output file name for the modified workbook
            string output = "SetDataDirection_result.xlsx";

            // Save the modified workbook to the specified file using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();
        
            //View the document
            FileViewer(output);
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
