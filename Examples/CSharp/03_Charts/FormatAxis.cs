using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Charts;

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
            this.label1.Text = "The sample demonstrates how to format axis for chart.";
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
            workbook.CreateEmptySheets(1);
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Demo";
            sheet.Range["A1"].Value = "Month";
            sheet.Range["A2"].Value = "Jan";
            sheet.Range["A3"].Value = "Feb";
            sheet.Range["A4"].Value = "Mar";
            sheet.Range["A5"].Value = "Apr";
            sheet.Range["A6"].Value = "May";
            sheet.Range["A7"].Value = "Jun";
            sheet.Range["A8"].Value = "Jul";
            sheet.Range["A9"].Value = "Aug";
            sheet.Range["B1"].Value = "Planned";
            sheet.Range["B2"].NumberValue = 38;
            sheet.Range["B3"].NumberValue = 47;
            sheet.Range["B4"].NumberValue = 39;
            sheet.Range["B5"].NumberValue = 36;
            sheet.Range["B6"].NumberValue = 27;
            sheet.Range["B7"].NumberValue = 25;
            sheet.Range["B8"].NumberValue = 36;
            sheet.Range["B9"].NumberValue = 48;

            Chart chart = sheet.Charts.Add(ExcelChartType.ColumnClustered);
            chart.DataRange = sheet.Range["B1:B9"];
            chart.SeriesDataFromRange = false;
            chart.PlotArea.Visible = false;
            chart.TopRow = 6;
            chart.BottomRow = 25;
            chart.LeftColumn = 2;
            chart.RightColumn = 9;
            chart.ChartTitle = "Chart with Customized Axis";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;
            Spire.Xls.Charts.ChartSerie cs1 = chart.Series[0];
            cs1.CategoryLabels = sheet.Range["A2:A9"];

            //format axis
            chart.PrimaryValueAxis.MajorUnit = 8;
            chart.PrimaryValueAxis.MinorUnit = 2;
            chart.PrimaryValueAxis.MaxValue = 50;
            chart.PrimaryValueAxis.MinValue = 0;
            chart.PrimaryValueAxis.IsReverseOrder = false;
            chart.PrimaryValueAxis.MajorTickMark = TickMarkType.TickMarkOutside;
            chart.PrimaryValueAxis.MinorTickMark = TickMarkType.TickMarkInside;
            chart.PrimaryValueAxis.TickLabelPosition = TickLabelPositionType.TickLabelPositionNextToAxis;
            chart.PrimaryValueAxis.CrossesAt = 0;
            //set NumberFormat
            chart.PrimaryValueAxis.NumberFormat = "$#,##0";
            chart.PrimaryValueAxis.IsSourceLinked = false;

            foreach (ChartSerie serie in chart.Series)
            {
                //format Series
                serie.DataPoints.DefaultDataPoint.DataFormat.Fill.FillType = ShapeFillType.SolidColor;
                serie.DataPoints.DefaultDataPoint.DataFormat.Fill.ForeColor = Color.Gray;
                serie.DataPoints.DefaultDataPoint.DataFormat.Fill.Transparency = 0.5;
                //format DataPoints
                serie.DataPoints[2].DataFormat.Fill.ForeColor = Color.Red;
            }



            workbook.SaveToFile("Result.xlsx", ExcelVersion.Version2010);
            ExcelDocViewer(workbook.FileName);
        }

        private void btnAbout_Click(object sender, System.EventArgs e)
        {
            Close();
        }


    }
}
