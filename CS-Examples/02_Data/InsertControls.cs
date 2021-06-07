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
            Workbook wb = new Workbook();
            wb.LoadFromFile(@"..\..\..\..\..\..\Data\InsertControls.xlsx");
            Worksheet ws = wb.Worksheets[0];

            //Add a textbox 
            ITextBoxShape textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100);
            textbox.Text = "Hello World";
            //Add a checkbox 
            ICheckBox cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100);
            cb.CheckState = Spire.Xls.CheckState.Checked;
            cb.Text = "Check Box 1";
            //Add a RadioButton 
            IRadioButton rb = ws.RadioButtons.Add(13, 2, 15, 100);
            rb.Text = "Option 1";

            //Add a combox
            IComboBoxShape cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100) as IComboBoxShape;
            cbx.ListFillRange = ws.Range["A41:A47"];

            wb.SaveToFile("Result.xlsx", ExcelVersion.Version2010);

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
