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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\MarkerDesigner2.xlsx");

            // Create a DataTable
            DataTable dt = new DataTable("data");

            //Define a field in it
            dt.Columns.Add(new DataColumn("value", typeof(string)));

            // Add three rows to it
            DataRow drName1 = dt.NewRow();
            DataRow drName2 = dt.NewRow();
            DataRow drName3 = dt.NewRow();

            drName1["value"] = "Text1";
            drName2["value"] = "Text2";
            drName3["value"] = "Text3";


            dt.Rows.Add(drName1);
            dt.Rows.Add(drName2);
            dt.Rows.Add(drName3);

            //Fill DataTable
            workbook.MarkerDesigner.AddDataTable("data", dt);
            workbook.MarkerDesigner.Apply();

            //Save the document
            string output = "SetDataDirection_result.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

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
