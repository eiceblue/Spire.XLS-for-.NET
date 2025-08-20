using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Text;
using System.IO;

namespace GetCellAddress
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

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            StringBuilder builder = new StringBuilder();

            //Get a cell range
            CellRange range = sheet.Range["A1:B5"];

            //Get address of range
            string address = range.RangeAddressLocal;
            builder.AppendLine("Address of range: " + address);

            //Get the cell count of range
            int count = range.CellsCount;
            builder.AppendLine("Cell count of range: " + count.ToString());

            //Get the address of the entire column of range
            string entireColAddress = range.EntireColumn.RangeAddressLocal;
            builder.AppendLine("Address of entire column of the range: " + entireColAddress);

            //Get the address of the entire row of range
            string entireRowAddress = range.EntireRow.RangeAddressLocal;
            builder.AppendLine("Address of entire row of the range " + entireRowAddress);
            
            //Write to txt file
            string output = "GetCellAddress_out.txt";
            File.WriteAllText(output, builder.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the txt file
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
