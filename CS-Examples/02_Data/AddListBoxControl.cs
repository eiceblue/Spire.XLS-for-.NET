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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ExcelSample_N1.xlsx");

            //Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Set text for cells 
            sheet.Range["A7"].Text = "Beijing";
            sheet.Range["A8"].Text = "New York";
            sheet.Range["A9"].Text = "ChengDu";
            sheet.Range["A10"].Text = "Paris";
            sheet.Range["A11"].Text = "Boston";
            sheet.Range["A12"].Text = "London";

            sheet.Range["C13"].Text = "City :";
            sheet.Range["C13"].Style.Font.IsBold = true;

            //Add listbox control
            IListBox listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80);
            listBox.SelectionType = SelectionType.Single;
            listBox.SelectedIndex = 2;
            listBox.Display3DShading = true;
            listBox.ListFillRange = sheet.Range["A7:A12"];

            //Save the document
            string output = "InsertListBoxControl_out.xlsx";
            workbook.SaveToFile(output, ExcelVersion.Version2013);

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
