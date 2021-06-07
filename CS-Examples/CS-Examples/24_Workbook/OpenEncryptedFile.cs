using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using System.Text;
using System.IO;

namespace OpenEncryptedFile
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //File path
            string filePath = @"..\..\..\..\..\..\Data\EncryptedFile.xlsx";

            //Create string builder
            StringBuilder builder = new StringBuilder();

            String[] passwords = new String[4] { "password1", "password2", "password3", "1234" };
            for (int i = 0; i < passwords.Length; i++)
            {
                try
                {
                    //Create a workbook
                    Workbook workbook = new Workbook();

                    //Open password
                    workbook.OpenPassword = passwords[i];

                    //Load the document
                    workbook.LoadFromFile(filePath);

                    builder.AppendLine("Password = " + passwords[i] + " is correct."+" The encrypted Excel file opened successfully!");
                }
                catch (Exception ex)
                {
                    builder.AppendLine("Password = " + passwords[i] + "  is not correct");
                }
            }

            //Save to txt file
            string result = "OpenEncryptedFile_out.txt";
            File.WriteAllText(result,builder.ToString());

            //Launch the file
            ExcelDocViewer(result);
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
