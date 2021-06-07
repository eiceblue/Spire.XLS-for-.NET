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
            //Create a workbook
            Workbook workbook = new Workbook();

            //Get the default first  worksheet
            Worksheet sheet = workbook.Worksheets[0];

            //Insert a TextBox
            sheet.Range["A2"].Text = "Name£º";
            ITextBoxShape textBox = sheet.TextBoxes.AddTextBox(2, 2, 18, 65);

            //Set the name 
            textBox.Name = "FirstTextBox";

            //Set string text for TextBox 
            textBox.Text = "Spire.XLS for .NET is a professional Excel .NET component that can be used to any type of .NET 2.0, 3.5, 4.0 or 4.5 framework application, both ASP.NET web sites and Windows Forms application.";

            //Get the TextBox by the name
            ITextBoxShape FindTextBox = sheet.TextBoxes["FirstTextBox"];

            //Get the TextBox text 
            string text = FindTextBox.Text;

            //Create StringBuilder to save 
            StringBuilder content = new StringBuilder();

            //Set string format for displaying
            string result = string.Format("The text of \"" + textBox.Name+"\" is :"+ text);

            //Add result string to StringBuilder
            content.AppendLine(result);

            //String for output file 
            String outputFile = "Output.txt";

            //Save them to a txt file
            File.WriteAllText(outputFile, content.ToString());

            //Launching the output file.
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
