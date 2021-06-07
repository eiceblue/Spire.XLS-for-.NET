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

            //Save the Excel file
            string result = "ImportDataFromArrayList_out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2013);

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
