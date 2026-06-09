using Spire.Xls;
using Spire.Xls.Core.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ToMarkdownExportOptions

{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            //Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToMarkdownExportOptions.xlsx");

            // Create export options for Markdown format
            MarkdownOptions options = new MarkdownOptions();

            // Set whether to save images with relative paths
            options.SavePicInRelativePath = true;

            // Set whether to save hyperlinks as Markdown reference format
            options.SaveHyperlinkAsRef = true;

            // Save the workbook as Markdown with the specified options
            string result = "ToMarkdownExportOptions_out.md";
            workbook.SaveToMarkdown(result, options);
            
            // Dispose of the workbook object to release resources
            workbook.Dispose();

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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
