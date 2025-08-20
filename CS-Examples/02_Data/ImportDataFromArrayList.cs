using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

namespace ImportDataFromArrayList
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //Create a workbook
            Workbook workbook = new Workbook();

            //Create an empty worksheet
            workbook.CreateEmptySheets(1);

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Create an ArrayList object
            ArrayList list = new ArrayList();

            //Add strings in list
            list.Add("Spire.Doc for .NET");
            list.Add("Spire.XLS for .NET");
            list.Add("Spire.PDF for .NET");
            list.Add("Spire.Presentation for .NET");

            //Insert array list in worksheet 
            sheet.InsertArrayList(list, 1, 1, true);

            // Specify the name for the resulting Excel file
            string result = "ImportDataFromArrayList_out.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
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
