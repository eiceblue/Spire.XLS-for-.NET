using System;
using System.Windows.Forms;
using Spire.Xls;
using System.IO;

namespace ObtainActiveSelectionRange
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Load an existing workbook from a file
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ObtainActiveSelectionRange.xlsx");

            // Get the first sheet
            Worksheet worksheet = workbook.Worksheets[0];

            string information = null;

            // Get the information of the active selection range
            foreach (CellRange range in worksheet.ActiveSelectionRange)
            {
                information += "RangeAddressLocal:" + range.RangeAddressLocal + "\r\n";
                information += "ColumnCount:" + range.ColumnCount + "\r\n";
                information += "ColumnWidth:" + range.ColumnWidth + "\r\n";
                information += "Column:" + range.Column + "\r\n";
                information += "RowCount:" + range.RowCount + "\r\n";
                information += "RowHeight:" + range.RowHeight + "\r\n";
                information += "Row:" + range.Row + "\r\n";
            }

            // Specify the output file name for the result
            string result = "ObtainActiveSelectionRange_result.txt";

            // Write the content of the information to the result file
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
