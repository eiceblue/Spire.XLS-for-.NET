using Spire.Xls;
using System;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CheckAutoFitRowOrColumn
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

            StringBuilder result = new StringBuilder();
            //Load an excel file
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\CheckAutoFitRowsAndColumns.xlsx");
            
            // Check if the second row has auto-fit row height set
            bool isRowAutofit = workbook.Worksheets[0].GetRowIsAutoFit(2);
            if (isRowAutofit)
            {
                result.AppendLine("The second row is auto fit row height.");
            }
            else {
                result.AppendLine("The second row is not auto fit row height.");
            }

            // Check if the second column has auto-fit column width set
            bool isColAutofit = workbook.Worksheets[0].GetColumnIsAutoFit(2);
            if (isColAutofit)
            {
                result.AppendLine("The second column is auto fit column width.");
            }
            else
            {
                result.AppendLine("The second column is not auto fit column width.");
            }

            // Save the result to a text file
            File.WriteAllText("CheckAutoFitRowOrColumn_result.txt", result.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            FileViewer("CheckAutoFitRowOrColumn_result.txt");
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
