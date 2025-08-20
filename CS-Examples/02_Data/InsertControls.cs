using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace InsertControls
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
            Workbook wb = new Workbook();

            // Load the workbook from file
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\InsertControls.xlsx");

            // Get the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Add a textbox at position (9, 2) with width 25 and height 100
            ITextBoxShape textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100);
            // Set the text for the textbox
            textbox.Text = "Hello World"; 

            // Add a checkbox at position (11, 2) with width 15 and height 100
            ICheckBox cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100);
            // Set the checkbox state to checked
            cb.CheckState = Spire.Xls.CheckState.Checked;
            // Set the text for the checkbox
            cb.Text = "Check Box 1"; 

            // Add a radio button at position (13, 2) with width 15 and height 100
            IRadioButton rb = ws.RadioButtons.Add(13, 2, 15, 100);
            // Set the text for the radio button
            rb.Text = "Option 1"; 

            // Add a combobox at position (15, 2) with width 15 and height 100
            IComboBoxShape cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100) as IComboBoxShape;
            // Set the range of options for the combobox
            cbx.ListFillRange = ws.Range["A41:A47"]; 

            // Save the workbook to file "Result.xlsx" in Excel 2010 format
            wb.SaveToFile("Result.xlsx", ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            wb.Dispose();

            //Launch the Excel file
            ExcelDocViewer("Result.xlsx");
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
