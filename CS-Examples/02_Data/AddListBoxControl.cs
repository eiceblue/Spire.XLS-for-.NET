using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace AddListBoxControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Load the Excel document from the specified file path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            // Get the first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Set text for cells A7 to A12
            sheet.Range["A7"].Text = "Beijing";
            sheet.Range["A8"].Text = "New York";
            sheet.Range["A9"].Text = "ChengDu";
            sheet.Range["A10"].Text = "Paris";
            sheet.Range["A11"].Text = "Boston";
            sheet.Range["A12"].Text = "London";

            // Set text and formatting for cell C13
            sheet.Range["C13"].Text = "City :";
            sheet.Range["C13"].Style.Font.IsBold = true;

            // Add a listbox control to the worksheet
            IListBox listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80);
            // Set the selection type to single (allows only one item to be selected)
            listBox.SelectionType = SelectionType.Single;
            // Set the initially selected index in the listbox
            listBox.SelectedIndex = 2;
            // Enable 3D shading for the listbox
            listBox.Display3DShading = true;
            // Specify the range to populate the listbox with data
            listBox.ListFillRange = sheet.Range["A7:A12"];

            //Specify the filename for the resulting Excel file
            string output = "InsertListBoxControl_out.xlsx";

            // Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            //Launch the Excel file
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
