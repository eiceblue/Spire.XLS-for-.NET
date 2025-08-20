using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using Spire.Xls;
using Spire.Xls.Charts;
using System.Collections.Generic;
using Spire.Xls.Core.Spreadsheet.PivotTables;
using System.Drawing.Imaging;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet;
using System.Text;

namespace GetTextBoxByName
{
	public partial class Form1 : Form
	{
        public Form1()
        {
            InitializeComponent();
        }
		private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook object
            Workbook workbook = new Workbook();

            // Get the default first worksheet from the workbook
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a TextBox at cell A2
            sheet.Range["A2"].Text = "Name£º";
            ITextBoxShape textBox = sheet.TextBoxes.AddTextBox(2, 2, 18, 65);

            // Set the name of the TextBox
            textBox.Name = "FirstTextBox";

            // Set the text for the TextBox
            textBox.Text = "Spire.XLS for .NET is a professional Excel .NET component that can be used in any type of .NET 2.0, 3.5, 4.0 or 4.5 framework application, both ASP.NET web sites and Windows Forms application.";

            // Get the TextBox by its name
            ITextBoxShape FindTextBox = sheet.TextBoxes["FirstTextBox"];

            // Get the text content of the TextBox
            string text = FindTextBox.Text;

            // Create a StringBuilder object to save the result
            StringBuilder content = new StringBuilder();

            // Format and store the result string
            string result = string.Format("The text of \"{0}\" is: {1}", textBox.Name, text);
            content.AppendLine(result);

            // Specify the output file path
            string outputFile = "Output.txt";

            // Save the result to a text file
            File.WriteAllText(outputFile, content.ToString());

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the output file
            Viewer(outputFile);

        }
		private void Viewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}
        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

	}
}
