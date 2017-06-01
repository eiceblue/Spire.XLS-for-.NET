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
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
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
			this.btnAbout.TabIndex = 3;
			this.btnAbout.Text = "Close";
			this.btnAbout.Click += new System.EventHandler(this.btnAbout_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(134)));
			this.label1.Location = new System.Drawing.Point(16, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(417, 18);
			this.label1.TabIndex = 4;
			this.label1.Text = "The sample demonstrates how to write formulas into spreadsheet.";
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
			this.Text = "Formulas sample";
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

		private void ExcelDocViewer( string fileName )
		{
			try
			{
				System.Diagnostics.Process.Start(fileName);
			}
			catch{}
		}

		private void btnRun_Click(object sender, System.EventArgs e)
		{
			Workbook workbook = new Workbook();
			Worksheet sheet = workbook.Worksheets[0];

			int currentRow = 1;
			string currentFormula = string.Empty;

			sheet.SetColumnWidth(1, 32);
			sheet.SetColumnWidth(2, 16);
			sheet.SetColumnWidth(3, 16);

			sheet.Range[currentRow++,1].Value = "Examples of formulas :";
			sheet.Range[++currentRow,1].Value = "Test data:";

			CellRange range = sheet.Range["A1"];
			range.Style.Font.IsBold = true;
			range.Style.FillPattern = ExcelPatternType.Solid;
			range.Style.KnownColor = ExcelColors.LightGreen1;
			range.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Medium;
			
			//test data
			sheet.Range[currentRow,2].NumberValue = 7.3;
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

			currentFormula = "=SUBTOTAL(3,Sheet1!B2:E3)";
			sheet.Range[currentRow, 1].Text = currentFormula;
			sheet.Range[currentRow++, 2].Formula = currentFormula;
			workbook.SaveToFile("Sample.xls");
			ExcelDocViewer(workbook.FileName);
		}

		private void btnAbout_Click(object sender, System.EventArgs e)
		{
			Close();
		}


	}
}
