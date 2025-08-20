using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using Spire.Xls;

namespace ImportDataFromDataColumn
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

            //Import the two columns of the data table to worksheet
            DataColumn[] columns=new DataColumn[2]{dataTable.Columns[1],dataTable.Columns[2]};
            sheet.InsertDataColumns(columns, true, 1, 1);

            // Specify the name for the resulting Excel file
            string result = "ImportDataFromDataColumn_output.xlsx";

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
