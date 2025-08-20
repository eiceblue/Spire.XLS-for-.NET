using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Text;
using System.IO;
using Spire.Xls.Core;
using Spire.Xls.Core.Spreadsheet;

namespace GetsCommentInNameManager
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new Workbook object
            Workbook workbook = new Workbook();

            // Load the Excel file "GetNotesInformation.xlsx" from a specific path
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\GetNotesInformation.xlsx");

            // Access the NameRanges property of the workbook
            INameRanges nameManager = workbook.NameRanges;

            // Create a StringBuilder to store the result
            StringBuilder stringBuilder = new StringBuilder();

            // Iterate through each name in the NameRanges collection
            for (int i = 0; i < nameManager.Count; i++)
            {
                // Get the XlsName object at index i
                XlsName name = (XlsName)nameManager[i];

                // Append the name and comment value to the StringBuilder
                stringBuilder.Append("Name: " + name.Name + ", Comment: " + name.CommentValue + "\r\n");
            }

            // Write the result to a text file named "GetsCommentInNameManager_result.txt"
            File.WriteAllText("GetsCommentInNameManager_result.txt", stringBuilder.ToString());

            // Dispose of the workbook object
            workbook.Dispose();
            // Launch the file
            ExcelDocViewer("GetsCommentInNameManager_result.txt");
		}
		private void ExcelDocViewer( string fileName )
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
