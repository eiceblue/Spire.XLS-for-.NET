using Spire.Xls;
using System;
using System.Windows.Forms;

namespace Subtotal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\Subtotal.xlsx");

            Worksheet sheet = workbook.Worksheets[0];
            //Select data range
            CellRange range = sheet.Range["A1:B18"];
            //Subtotal selected data
            sheet.Subtotal(range, 0, new int[] {1}, SubtotalTypes.Sum, true, false, true);

            //Save to file
            String result = "Subtotal_Out.xlsx";
            workbook.SaveToFile(result, ExcelVersion.Version2010);
            
            //View the document
            FileViewer(result);
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
