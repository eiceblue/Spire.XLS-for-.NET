using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

using Spire.Xls;

namespace Spire.Xls.Sample
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
    {
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnAbout;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGrid dataGrid1;
        private Button btnRead;
        private Button btnCalculate;
        private ToolTip toolTip;
        private IContainer components;

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
            this.components = new System.ComponentModel.Container();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnAbout = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGrid1 = new System.Windows.Forms.DataGrid();
            this.btnRead = new System.Windows.Forms.Button();
            this.btnCalculate = new System.Windows.Forms.Button();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExport
            // 
            this.btnExport.Location = new System.Drawing.Point(360, 288);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(75, 23);
            this.btnExport.TabIndex = 2;
            this.btnExport.Text = "Export";
            this.toolTip.SetToolTip(this.btnExport, "Exports cells to datatable with calculating formulas");
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnAbout
            // 
            this.btnAbout.Location = new System.Drawing.Point(448, 288);
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Size = new System.Drawing.Size(75, 23);
            this.btnAbout.TabIndex = 3;
            this.btnAbout.Text = "Close";
            this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(16, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(528, 32);
            this.label1.TabIndex = 4;
            this.label1.Text = "The sample demonstrates how to calculate formulas and export data to datatable wi" +
                "th calculating formulas.";
            // 
            // dataGrid1
            // 
            this.dataGrid1.DataMember = "";
            this.dataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid1.Location = new System.Drawing.Point(16, 56);
            this.dataGrid1.Name = "dataGrid1";
            this.dataGrid1.ReadOnly = true;
            this.dataGrid1.Size = new System.Drawing.Size(512, 216);
            this.dataGrid1.TabIndex = 5;
            // 
            // btnRead
            // 
            this.btnRead.Location = new System.Drawing.Point(272, 288);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(75, 23);
            this.btnRead.TabIndex = 2;
            this.btnRead.Text = "Read";
            this.toolTip.SetToolTip(this.btnRead, "Calculates cells\' formulas and reads results");
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // btnCalculate
            // 
            this.btnCalculate.Location = new System.Drawing.Point(184, 288);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(75, 23);
            this.btnCalculate.TabIndex = 2;
            this.btnCalculate.Text = "Calculate";
            this.toolTip.SetToolTip(this.btnCalculate, "Calculates one special formula");
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);
            // 
            // toolTip
            // 
            this.toolTip.IsBalloon = true;
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
            this.ClientSize = new System.Drawing.Size(544, 325);
            this.Controls.Add(this.dataGrid1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAbout);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.btnRead);
            this.Controls.Add(this.btnExport);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Spire.XLS sample";
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
            this.ResumeLayout(false);

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

        private Stream WriteFormulas()
        {
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            int currentRow = 1;
            string currentFormula = string.Empty;

            sheet.SetColumnWidth(1, 32);
            sheet.SetColumnWidth(2, 16);
            sheet.SetColumnWidth(3, 16);

            sheet.Range[currentRow++, 1].Value = "Examples of formulas :";
            sheet.Range[++currentRow, 1].Value = "Test data:";

            CellRange range = sheet.Range["A1"];
            range.Style.Font.IsBold = true;
            range.Style.FillPattern = ExcelPatternType.Solid;
            range.Style.KnownColor = ExcelColors.LightGreen1;
            range.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium;

            //test data
            sheet.Range[currentRow, 2].NumberValue = 7.3;
            sheet.Range[currentRow, 3].NumberValue = 5; ;
            sheet.Range[currentRow, 4].NumberValue = 8.2;
            sheet.Range[currentRow, 5].NumberValue = 4;
            sheet.Range[currentRow, 6].NumberValue = 3;
            sheet.Range[currentRow, 7].NumberValue = 11.3;

            sheet.Range[++currentRow, 1].Value = "Formulas"; ;
            sheet.Range[currentRow, 2].Value = "Results";
            range = sheet.Range[currentRow, 1, currentRow, 2];
            //range.Value = "Formulas";
            range.Style.Font.IsBold = true;
            range.Style.KnownColor = ExcelColors.LightGreen1;
            range.Style.FillPattern = ExcelPatternType.Solid;
            range.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium;
            //str.
            currentFormula = "=\"hello\"";
            sheet.Range[++currentRow, 1].Text = "=\"hello\"";
            sheet.Range[currentRow, 2].Formula = currentFormula;
            sheet.Range[currentRow, 3].Formula = "=\"" + new string(new char[] { '\u4f60', '\u597d' }) + "\"";

            //int.
            currentFormula = "=300";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;

            // float
            currentFormula = "=3389.639421";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;

            //bool.
            currentFormula = "=false";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;

            currentFormula = "=1+2+3+4+5-6-7+8-9";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;

            currentFormula = "=33*3/4-2+10";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;


            // sheet reference
            currentFormula = "=Sheet1!$B$3";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;

            // sheet area reference
            currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;

            // Functions
            currentFormula = "=Count(3,5,8,10,2,34)";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;


            currentFormula = "=NOW()";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow, 2].Formula = currentFormula;
            sheet.Range[currentRow, 2].Style.NumberFormat = "yyyy-MM-DD";

            currentFormula = "=SECOND(11)";
            sheet.Range[++currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=MINUTE(12)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=MONTH(9)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=DAY(10)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=TIME(4,5,7)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=DATE(6,4,2)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=RAND()";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=HOUR(12)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=MOD(5,3)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=WEEKDAY(3)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=YEAR(23)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=NOT(true)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=OR(true)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=AND(TRUE)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=VALUE(30)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=LEN(\"world\")";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=MID(\"world\",4,2)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=ROUND(7,3)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=SIGN(4)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=INT(200)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=ABS(-1.21)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=LN(15)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=EXP(20)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=SQRT(40)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=PI()";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=COS(9)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=SIN(45)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=MAX(10,30)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=MIN(5,7)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=AVERAGE(12,45)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=SUM(18,29)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            currentFormula = "=IF(4,2,2)";
            sheet.Range[currentRow, 1].Text = currentFormula;
            sheet.Range[currentRow++, 2].Formula = currentFormula;

            MemoryStream buffer = new MemoryStream();
            workbook.SaveToStream(buffer);
            buffer.Position = 0;

            return buffer;
        }

		private void btnExport_Click(object sender, System.EventArgs e)
		{
            using (Stream buffer = this.WriteFormulas())
            {
                //load
                Workbook workbook = new Workbook();
                workbook.LoadFromStream(buffer);

                //calculate all cells
                workbook.CalculateAllValue();

                //export
                Worksheet sheet = workbook.Worksheets[0];
                this.dataGrid1.DataSource = sheet.ExportDataTable(sheet["A4:B46"], true, true);
            }
		}

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            using (Stream buffer = this.WriteFormulas())
            {
                //load
                Workbook workbook = new Workbook();
                workbook.LoadFromStream(buffer);

                //calculate formula
                Object b3 = workbook.CaculateFormulaValue("Sheet1!$B$3");
                Object c3 = workbook.CaculateFormulaValue("Sheet1!$C$3");
                String formula = "Sheet1!$B$3 + Sheet1!$C$3";
                Object value = workbook.CaculateFormulaValue(formula);
                String message
                    = String.Format("Sheet1!$B$3 = {0}, Sheet1!$C$3 = {1}, {2} = {3}",
                        b3, c3, formula, value);
                MessageBox.Show(message);
            }
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            using (Stream buffer = this.WriteFormulas())
            {
                //load
                Workbook workbook = new Workbook();
                workbook.LoadFromStream(buffer);

                //calculate all cells' formula
                workbook.CalculateAllValue();

                //read cells' value to data table
                Worksheet sheet = workbook.Worksheets[0];
                DataTable dataTable = new DataTable();
                dataTable.Columns.Add("Formulas", typeof(String));
                dataTable.Columns.Add("Results", typeof(Object));
                foreach (CellRange row in sheet["A5:B46"].Rows)
                {
                    String formula = row.Columns[1].Formula;
                    Object value = row.Columns[1].FormulaValue;
                    dataTable.Rows.Add(formula, value);
                }
                this.dataGrid1.DataSource = dataTable;
            }
        }

        private void btnAbout_Click(object sender, System.EventArgs e)
        {
            Close();
        }
	}
}
