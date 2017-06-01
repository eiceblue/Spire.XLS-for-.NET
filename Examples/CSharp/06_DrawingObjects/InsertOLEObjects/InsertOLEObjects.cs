using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core;

namespace Spire.Xls.Sample
{
    /// <summary>
    /// Summary description for Form1.
    /// </summary>
    public class Form1 : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.Label label1;
        /// <summary>
        /// Required designer variable.
        /// </summary
        private System.ComponentModel.Container components = null;

        public Form1()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();
            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnRun = new System.Windows.Forms.Button();
            this.btnAbout = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(278, 59);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(72, 23);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "Run";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Location = new System.Drawing.Point(356, 59);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(75, 23);
            this.btnAbout.TabIndex = 3;
            this.btnAbout.Text = "Close";
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(16, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(467, 14);
            this.label1.TabIndex = 4;
            this.label1.Text = "The sample demonstrates how to insert OLE Objects in an Excel workbook.";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(495, 95);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.btnRun);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Formulas sample";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.Run(new Form1());
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            //load Excel file
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Range["A1"].Text = "The sample demonstrates how to insert OLE Objects in an Excel workbook.";
            //insert OLE object
            string xlsFile = @"..\..\..\..\..\..\Data\MiscDataTable.xls";
            Image image = GenerateImage(xlsFile);
            IOleObject oleObject = ws.OleObjects.Add(xlsFile, image, OleLinkType.Embed);
            oleObject.Location = ws.Range["B4"];
            oleObject.ObjectType = OleObjectType.ExcelWorksheet;
            //save the file
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer(workbook.FileName);
        }

        private void btnAbout_Click(object sender, System.EventArgs e)
        {
            Close();
        }
        private Image GenerateImage(string fileName)
        {
            Workbook book = new Workbook();
            book.LoadFromFile(fileName);
            book.Worksheets[0].PageSetup.LeftMargin = 0;
            book.Worksheets[0].PageSetup.RightMargin = 0;
            book.Worksheets[0].PageSetup.TopMargin = 0;
            book.Worksheets[0].PageSetup.BottomMargin = 0;
            return book.Worksheets[0].ToImage(1, 1, 19, 5);
        }
        private void ExcelDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }


    }
}
