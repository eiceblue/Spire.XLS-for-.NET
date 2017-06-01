using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;

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
            this.btnRun.Location = new System.Drawing.Point(448, 56);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(72, 23);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "Run";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Location = new System.Drawing.Point(528, 56);
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
            this.label1.Size = new System.Drawing.Size(553, 28);
            this.label1.TabIndex = 4;
            this.label1.Text = "The sample demonstrates how to insert SparkLine into an excel workbook.";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(616, 94);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.btnRun);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Spire.XLS sample";
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

        private void ExcelDocViewer(string fileName)
        {
            try
            {
                System.Diagnostics.Process.Start(fileName);
            }
            catch { }
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.Version = ExcelVersion.Version2010;
            workbook.CreateEmptySheets(1);

            Worksheet sheet = workbook.Worksheets[0];

            //Country
            sheet.Range["A1"].Value = "Country";
            sheet.Range["A2"].Value = "Cuba";
            sheet.Range["A3"].Value = "Mexico";
            sheet.Range["A4"].Value = "France";
            sheet.Range["A5"].Value = "German";

            //Jun
            sheet.Range["B1"].Value = "Jun";
            sheet.Range["B2"].NumberValue = 0.23;
            sheet.Range["B3"].NumberValue = 0.37;
            sheet.Range["B4"].NumberValue = 0.15;
            sheet.Range["B5"].NumberValue = 0.25;

            //Jul
            sheet.Range["C1"].Value = "Jul";
            sheet.Range["C2"].NumberValue = 0.1;
            sheet.Range["C3"].NumberValue = 0.35;
            sheet.Range["C4"].NumberValue = 0.22;
            sheet.Range["C5"].NumberValue = 0.33;


            //Aug
            sheet.Range["D1"].Value = "Aug";
            sheet.Range["D2"].NumberValue = 0.14;
            sheet.Range["D3"].NumberValue = 0.36;
            sheet.Range["D4"].NumberValue = 0.25;
            sheet.Range["D5"].NumberValue = 0.25;


            //Aug
            sheet.Range["E1"].Value = "Sep";
            sheet.Range["E2"].NumberValue = 0.17;
            sheet.Range["E3"].NumberValue = 0.28;
            sheet.Range["E4"].NumberValue = 0.39;
            sheet.Range["E5"].NumberValue = 0.32;

            //Style
            sheet.Range["A1:E1"].Style.Font.IsBold = true;
            sheet.Range["A2:E2"].Style.KnownColor = ExcelColors.LightYellow;
            sheet.Range["A3:E3"].Style.KnownColor = ExcelColors.LightGreen1;
            sheet.Range["A4:E4"].Style.KnownColor = ExcelColors.LightOrange;
            sheet.Range["A5:E5"].Style.KnownColor = ExcelColors.LightTurquoise;

            //Border
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeTop].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeBottom].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeLeft].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeRight].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;

            sheet.Range["B2:D5"].Style.NumberFormatIndex = 9;

            SparklineGroup sparklineGroup
                = sheet.SparklineGroups.AddGroup(SparklineType.Line);
            SparklineCollection sparklines = sparklineGroup.Add();
            sparklines.Add(sheet["B2:E2"], sheet["F2"]);
            sparklines.Add(sheet["B3:E3"], sheet["F3"]);
            sparklines.Add(sheet["B4:E4"], sheet["F4"]);
            sparklines.Add(sheet["B5:E5"], sheet["F5"]);

            workbook.SaveToFile("Sample.xlsx");

            ExcelDocViewer(workbook.FileName);
        }

        private void btnAbout_Click(object sender, System.EventArgs e)
        {
            Close();
        }


    }
}
