using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;

namespace ImportDataFromDataTable
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

            //Create a DataTable object 
            DataTable dataTable = new DataTable("Customer");
            dataTable.Columns.Add("No", typeof(Int32));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("City", typeof(string));

            //Create rows and add data
            DataRow dr = dataTable.NewRow();
            dr[0] = 1;
            dr[1] = "Tom";
            dr[2] = "New York";
            dataTable.Rows.Add(dr);
            dr = dataTable.NewRow();
            dr[0] = 2;
            dr[1] = "Jerry";
            dr[2] = "China";
            dataTable.Rows.Add(dr);
            dr = dataTable.NewRow();
            dr[0] = 3;
            dr[1] = "Dive Time";
            dr[2] = "Berkely";
            dataTable.Rows.Add(dr);
            dr = dataTable.NewRow();
            dr[0] = 4;
            dr[1] = "Amor Aqua";
            dr[2] = "Florida";
            dataTable.Rows.Add(dr);

            //Import datatable in worksheet
            sheet.InsertDataTable(dataTable, true, 1, 1);

            //Save the Excel file
            string result = "ImportDataFromDataTable_output.xlsx";
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
