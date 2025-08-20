using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace CopyShapes
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Create line shape
            var line = sheet.TypedLines.AddLine();
            line.Top = 50;
            line.Left = 30;
            line.Width = 30;
            line.Height = 50;
            line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowDiamond;
            line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow;

            // Get the second worksheet
            Worksheet CopyShapes = workbook.Worksheets[1];

            // Copy the line into other sheet
            CopyShapes.TypedLines.AddCopy(line);

            // Create a button and then copy into other sheet
            var button = sheet.TypedRadioButtons.Add( 5, 5, 20, 20);
            CopyShapes.TypedRadioButtons.AddCopy(button);

            // Create a textbox and then copy into other sheet
            var textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100);
            CopyShapes.TypedTextBoxes.AddCopy(textbox);

            // Create a checkbox and then copy into other sheet
            var checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20);
            CopyShapes.TypedCheckBoxes.AddCopy(checkbox);

            // Create a comboboxes and then copy into other sheet
            sheet.Range["A14"].Value = "1";
            sheet.Range["A15"].Value = "2";
            var ComboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30);
            ComboBoxes.ListFillRange = sheet.Range["A14:A15"];
            CopyShapes.TypedComboBoxes.AddCopy(ComboBoxes);

            // Save the file
            workbook.SaveToFile("CopyShapes.xlsx",ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the document
            FileViewer("CopyShapes.xlsx");
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
