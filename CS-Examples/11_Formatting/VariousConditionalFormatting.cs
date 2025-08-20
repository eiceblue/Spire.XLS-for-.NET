using System;
using System.Data.OleDb;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using Spire.Xls;
using Spire.Xls.Core.Spreadsheet.ConditionalFormatting;
using Spire.Xls.Core.Spreadsheet.Collections;
using Spire.Xls.Core;

namespace VariousConditionalFormatting
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, System.EventArgs e)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Create a blank sheet
            workbook.CreateEmptySheets(1);

            // Get the first sheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add conditional formatting for sheet
            AddConditionalFormattingForNewSheet(sheet);

            // Specify the name for the resulting Excel file
            String result = "VariousConditionalFormatting_result.xlsx";

            // Save the workbook to the specified file in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010);

            // Dispose of the workbook object to release resources
            workbook.Dispose();

            // Launch the file
            ExcelDocViewer(result);
        }
        private void AddConditionalFormattingForNewSheet(Worksheet sheet)
        {
            AddDefaultIconSet(sheet);
            AddIconSet2(sheet);
            AddIconSet3(sheet);
            AddIconSet4(sheet);
            AddIconSet5(sheet);
            AddIconSet6(sheet);
            AddIconSet7(sheet);
            AddIconSet8(sheet);
            AddIconSet9(sheet);
            AddIconSet10(sheet);
            AddIconSet11(sheet);
            AddIconSet12(sheet);
            AddIconSet13(sheet);
            AddIconSet14(sheet);
            AddIconSet15(sheet);
            AddIconSet16(sheet);
            AddIconSet17(sheet);
            AddIconSet18(sheet);
            AddDefaultColorScale(sheet);
            Add3ColorScale(sheet);
            Add2ColorScale(sheet);
            AddAboveAverage(sheet);
            AddAboveAverage2(sheet);
            AddAboveAverage3(sheet);
            AddTop10_1(sheet);
            AddTop10_2(sheet);
            AddTop10_3(sheet);
            AddTop10_4(sheet);
            AddDataBar1(sheet);
            AddDataBar2(sheet);
            AddContainsText(sheet);
            AddNotContainsText(sheet);
            AddContainsBlank(sheet);
            AddNotContainsBlank(sheet);
            AddBeginWith(sheet);
            AddEndWith(sheet);
            AddContainsError(sheet);
            AddNotContainsError(sheet);
            AddDuplicate(sheet);
            AddUnique(sheet);
            AddTimePeriod_1(sheet);
            AddTimePeriod_2(sheet);
            AddTimePeriod_3(sheet);
            AddTimePeriod_4(sheet);
            AddTimePeriod_5(sheet);
            AddTimePeriod_6(sheet);
            AddTimePeriod_7(sheet);
            AddTimePeriod_8(sheet);
            AddTimePeriod_9(sheet);
            AddTimePeriod_10(sheet);
            sheet.AllocatedRange.ColumnWidth = 15;
            sheet.AllocatedRange.AutoFitRows();
        }
        //This method implements the IconSet conditional formatting type with ThreeArrows colored attribute.
        private void AddIconSet2(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M1:O2"]);
            sheet.Range["M1:O2"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M1:O2"].Style.Color = Color.AliceBlue;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeArrows;
            sheet.Range["M1"].Text = "ThreeArrows";
            sheet.Range["N1"].NumberValue = 15;
            sheet.Range["O1"].NumberValue = 18;
            sheet.Range["M2"].NumberValue = 14;
            sheet.Range["N2"].NumberValue = 17;
            sheet.Range["O2"].NumberValue = 20;
        }
        //This method implements the IconSet conditional formatting type with FourArrows colored attribute.
        private void AddIconSet3(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M3:O4"]);
            sheet.Range["M3:O4"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M3:O4"].Style.Color = Color.AntiqueWhite;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FourArrows;
            sheet.Range["M3"].Text = "FourArrows";
            sheet.Range["N3"].NumberValue = 17;
            sheet.Range["O3"].NumberValue = 20;
            sheet.Range["M4"].NumberValue = 16;
            sheet.Range["N4"].NumberValue = 19;
            sheet.Range["O4"].NumberValue = 22;
        }
        //This method implements the IconSet conditional formatting type with FiveArrows colored attribute.
        private void AddIconSet4(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M5:O6"]);
            sheet.Range["M5:O6"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M5:O6"].Style.Color = Color.Aqua;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FiveArrows;
            sheet.Range["M5"].Text = "FiveArrows";
            sheet.Range["N5"].NumberValue = 17;
            sheet.Range["O5"].NumberValue = 20;
            sheet.Range["M6"].NumberValue = 16;
            sheet.Range["N6"].NumberValue = 19;
            sheet.Range["O6"].NumberValue = 22;
        }
        //This method implements the IconSet conditional formatting type with ThreeArrowsGray colored attribute.
        private void AddIconSet5(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M7:O8"]);
            sheet.Range["M7:O8"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M7:O8"].Style.Color = Color.Aquamarine;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeArrowsGray;
            sheet.Range["M7"].Text = "ThreeArrowsGray";
            sheet.Range["N7"].NumberValue = 21;
            sheet.Range["O7"].NumberValue = 24;
            sheet.Range["M8"].NumberValue = 20;
            sheet.Range["N8"].NumberValue = 23;
            sheet.Range["O8"].NumberValue = 26;
        }
        //This method implements the IconSet conditional formatting type with FourArrowsGray colored attribute.
        private void AddIconSet6(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M9:O10"]);
            sheet.Range["M9:O10"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M9:O10"].Style.Color = Color.Azure;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FourArrowsGray;
            sheet.Range["M9"].Text = "FourArrowsGray";
            sheet.Range["N9"].NumberValue = 23;
            sheet.Range["O9"].NumberValue = 26;
            sheet.Range["M10"].NumberValue = 22;
            sheet.Range["N10"].NumberValue = 25;
            sheet.Range["O10"].NumberValue = 28;
        }
        //This method implements the IconSet conditional formatting type with FiveArrowsGray colored attribute.
        private void AddIconSet7(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M11:O12"]);
            sheet.Range["M11:O12"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M11:O12"].Style.Color = Color.Beige;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FiveArrowsGray;
            sheet.Range["M11"].Text = "FiveArrowsGray";
            sheet.Range["N11"].NumberValue = 25;
            sheet.Range["O11"].NumberValue = 28;
            sheet.Range["M12"].NumberValue = 24;
            sheet.Range["N12"].NumberValue = 27;
            sheet.Range["O12"].NumberValue = 30;
        }
        //This method implements the IconSet conditional formatting type with ThreeFlags attribute.
        private void AddIconSet8(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M13:O14"]);
            sheet.Range["M13:O14"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M13:O14"].Style.Color = Color.Bisque;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeFlags;
            sheet.Range["M13"].Text = "ThreeFlags";
            sheet.Range["N13"].NumberValue = 27;
            sheet.Range["O13"].NumberValue = 30;
            sheet.Range["M14"].NumberValue = 26;
            sheet.Range["N14"].NumberValue = 29;
            sheet.Range["O14"].NumberValue = 32;
        }
        //This method implements the IconSet conditional formatting type with FiveQuarters attribute.
        private void AddIconSet9(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M15:O16"]);
            sheet.Range["M15:O16"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M15:O16"].Style.Color = Color.BlanchedAlmond;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FiveQuarters;
            sheet.Range["M15"].Text = "FiveQuarters";
            sheet.Range["N15"].NumberValue = 29;
            sheet.Range["O15"].NumberValue = 32;
            sheet.Range["M16"].NumberValue = 28;
            sheet.Range["N16"].NumberValue = 31;
            sheet.Range["O16"].NumberValue = 34;
        }
        //This method implements the IconSet conditional formatting type with FourRating attribute.
        private void AddIconSet10(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M17:O18"]);
            sheet.Range["M17:O18"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M17:O18"].Style.Color = Color.LightBlue;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FourRating;
            sheet.Range["M17"].Text = "FourRating";
            sheet.Range["N17"].NumberValue = 31;
            sheet.Range["O17"].NumberValue = 34;
            sheet.Range["M18"].NumberValue = 30;
            sheet.Range["N18"].NumberValue = 33;
            sheet.Range["O18"].NumberValue = 36;
        }
        //This method implements the IconSet conditional formatting type with FiveRating attribute.
        private void AddIconSet11(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M19:O20"]);
            sheet.Range["M19:O20"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M19:O20"].Style.Color = Color.BlueViolet;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FiveRating;
            sheet.Range["M19"].Text = "FiveRating";
            sheet.Range["N19"].NumberValue = 33;
            sheet.Range["O19"].NumberValue = 36;
            sheet.Range["M20"].NumberValue = 32;
            sheet.Range["N20"].NumberValue = 35;
            sheet.Range["O20"].NumberValue = 38;
        }
        //This method implements the IconSet conditional formatting type with FourRedToBlack attribute.
        private void AddIconSet12(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M21:O22"]);
            sheet.Range["M21:O22"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M21:O22"].Style.Color = Color.Brown;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FourRedToBlack;
            sheet.Range["M21"].Text = "FourRedToBlack";
            sheet.Range["N21"].NumberValue = 35;
            sheet.Range["O21"].NumberValue = 38;
            sheet.Range["M22"].NumberValue = 34;
            sheet.Range["N22"].NumberValue = 37;
            sheet.Range["O22"].NumberValue = 40;
        }
        //This method implements the IconSet conditional formatting type with ThreeSigns attribute.
        private void AddIconSet13(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M23:O24"]);
            sheet.Range["M23:O24"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M23:O24"].Style.Color = Color.BurlyWood;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeSigns;
            sheet.Range["M23"].Text = "ThreeSigns";
            sheet.Range["N23"].NumberValue = 37;
            sheet.Range["O23"].NumberValue = 40;
            sheet.Range["M24"].NumberValue = 36;
            sheet.Range["N24"].NumberValue = 39;
            sheet.Range["O24"].NumberValue = 42;
        }
        //This method implements the IconSet conditional formatting type with ThreeSymbols attribute.
        private void AddIconSet14(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M25:O26"]);
            sheet.Range["M25:O26"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M25:O26"].Style.Color = Color.CadetBlue;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeSymbols;
            sheet.Range["M25"].Text = "ThreeSymbols";
            sheet.Range["N25"].NumberValue = 39;
            sheet.Range["O25"].NumberValue = 42;
            sheet.Range["M26"].NumberValue = 38;
            sheet.Range["N26"].NumberValue = 41;
            sheet.Range["O26"].NumberValue = 44;
        }
        //This method implements the IconSet conditional formatting type with ThreeSymbols2 attribute.
        private void AddIconSet15(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M27:O28"]);
            sheet.Range["M27:O28"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M27:O28"].Style.Color = Color.Chartreuse;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeSymbols2;
            sheet.Range["M27"].Text = "ThreeSymbols2";
            sheet.Range["N27"].NumberValue = 41;
            sheet.Range["O27"].NumberValue = 44;
            sheet.Range["M28"].NumberValue = 40;
            sheet.Range["N28"].NumberValue = 43;
            sheet.Range["O28"].NumberValue = 46;
        }
        //This method implements the IconSet conditional formatting type with ThreeTrafficLights1 attribute.
        private void AddIconSet16(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M29:O30"]);
            sheet.Range["M29:O30"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M29:O30"].Style.Color = Color.Chocolate;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeTrafficLights1;
            sheet.Range["M29"].Text = "ThreeTrafficLights1";
            sheet.Range["N29"].NumberValue = 43;
            sheet.Range["O29"].NumberValue = 46;
            sheet.Range["M30"].NumberValue = 42;
            sheet.Range["N30"].NumberValue = 45;
            sheet.Range["O30"].NumberValue = 48;
        }
        //This method implements the IconSet conditional formatting type with ThreeTrafficLights2 attribute.
        private void AddIconSet17(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M31:O32"]);
            sheet.Range["M31:O32"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M31:O32"].Style.Color = Color.Coral;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.ThreeTrafficLights2;
            sheet.Range["M31"].Text = "ThreeTrafficLights2";
            sheet.Range["N31"].NumberValue = 45;
            sheet.Range["O31"].NumberValue = 48;
            sheet.Range["M32"].NumberValue = 44;
            sheet.Range["N32"].NumberValue = 47;
            sheet.Range["O32"].NumberValue = 50;
        }
        //This method implements the IconSet conditional formatting type with FourTrafficLights attribute.
        private void AddIconSet18(Worksheet sheet)
        {
            XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
            xcfs.AddRange(sheet.Range["M33:O35"]);
            sheet.Range["M33:O35"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["M33:O35"].Style.Color = Color.CornflowerBlue;
            IConditionalFormat cf = xcfs.AddCondition();
            cf.FormatType = ConditionalFormatType.IconSet;
            cf.IconSet.IconSetType = IconSetType.FourTrafficLights;
            sheet.Range["M33"].Text = "FourTrafficLights";
            sheet.Range["N33"].NumberValue = 48;
            sheet.Range["O33"].NumberValue = 52;
            sheet.Range["M34"].NumberValue = 46;
            sheet.Range["N34"].NumberValue = 50;
            sheet.Range["O34"].NumberValue = 54;
            sheet.Range["M35"].NumberValue = 48;
            sheet.Range["N35"].NumberValue = 52;
            sheet.Range["O35"].NumberValue = 56;
        }
        //This method implements the TimePeriod conditional formatting type with Yesterday attribute.
         private void AddTimePeriod_10(Worksheet sheet)
        {
            XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
            conds.AddRange(sheet.Range["I19:K20"]);
            sheet.Range["I19:K20"].Style.FillPattern = ExcelPatternType.Solid;
            sheet.Range["I19:K20"].Style.Color = Color.MediumSeaGreen;
            IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.Yesterday);
            cf.FillPattern = ExcelPatternType.Solid;
            cf.BackColor = Color.Pink;
            CellRange c = sheet.Range["I19"];
            c.Value2 = DateTime.Now.AddDays(-2).Date;

            c = sheet.Range["J19"];
            c.Value2 = DateTime.Now.AddDays(-1).Date;
          
            c = sheet.Range["K19"];
            c.Value2 = DateTime.Now.Date;
  
            c = sheet.Range["I20"];
            c.Text = "Yesterday";

            c = sheet.Range["J20"];
            c.Value2 = DateTime.Now.AddDays(1).Date;

            c = sheet.Range["K20"];
            c.Value2 = DateTime.Now.AddDays(2).Date;
        }
         //This method implements the TimePeriod conditional formatting type with Tomorrow attribute.
         private void AddTimePeriod_9(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I17:K18"]);
             sheet.Range["I17:K18"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I17:K18"].Style.Color = Color.MediumPurple;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.Tomorrow);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I17"];
             c.Value2 = DateTime.Now.AddDays(-2).Date;

             c = sheet.Range["J17"];
             c.Value2 = DateTime.Now.AddDays(-1).Date;

             c = sheet.Range["K17"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I18"];
             c.Text = "Tomorrow";

             c = sheet.Range["J18"];
             c.Value2 = DateTime.Now.AddDays(1).Date;

             c = sheet.Range["K18"];
             c.Value2 = DateTime.Now.AddDays(2).Date;
         }
         //This method implements the TimePeriod conditional formatting type with ThisWeek attribute.
         private void AddTimePeriod_8(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I15:K16"]);
             sheet.Range["I15:K16"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I15:K16"].Style.Color = Color.MediumOrchid;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.ThisWeek);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I15"];
             c.Value2 = DateTime.Now.AddDays(-2).Date;

             c = sheet.Range["J15"];
             c.Value2 = DateTime.Now.AddDays(-1).Date;

             c = sheet.Range["K15"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I16"];
             c.Text = "ThisWeek";

             c = sheet.Range["J16"];
             c.Value2 = DateTime.Now.AddDays(2).Date;

             c = sheet.Range["K16"];
             c.Value2 = DateTime.Now.AddDays(3).Date;
         }
         //This method implements the TimePeriod conditional formatting type with ThisMonth attribute.
         private void AddTimePeriod_7(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I13:K14"]);
             sheet.Range["I13:K14"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I13:K14"].Style.Color = Color.MediumBlue;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.ThisMonth);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I13"];
             c.Value2 = DateTime.Now.AddMonths(-1).Date;

             c = sheet.Range["J13"];
             c.Value2 = DateTime.Now.AddDays(-1).Date;

             c = sheet.Range["K13"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I14"];
             c.Text = "ThisMonth";

             c = sheet.Range["J14"];
             c.Value2 = DateTime.Now.AddMonths(1).Date;

             c = sheet.Range["K14"];
             c.Value2 = DateTime.Now.AddMonths(2).Date;
         }
         //This method implements the TimePeriod conditional formatting type with NextWeek attribute.
         private void AddTimePeriod_6(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I11:K12"]);
             sheet.Range["I11:K12"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I11:K12"].Style.Color = Color.MediumAquamarine;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.NextWeek);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I11"];
             c.Value2 = DateTime.Now.AddDays(-3).Date;

             c = sheet.Range["J11"];
             c.Value2 = DateTime.Now.AddDays(-2).Date;

             c = sheet.Range["K11"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I12"];
             c.Text = "NextWeek";

             c = sheet.Range["J12"];
             c.Value2 = DateTime.Now.AddDays(3).Date;

             c = sheet.Range["K12"];
             c.Value2 = DateTime.Now.AddMonths(4).Date;
         }
         //This method implements the TimePeriod conditional formatting type with NextMonth attribute.
         private void AddTimePeriod_5(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I9:K10"]);
             sheet.Range["I9:K10"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I9:K10"].Style.Color = Color.Maroon;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.NextMonth);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I9"];
             c.Value2 = DateTime.Now.AddDays(-3).Date;

             c = sheet.Range["J9"];
             c.Value2 = DateTime.Now.AddMonths(-1).Date;

             c = sheet.Range["K9"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I10"];
             c.Text = "NextMonth";

             c = sheet.Range["J10"];
             c.Value2 = DateTime.Now.AddMonths(1).Date;

             c = sheet.Range["K10"];
             c.Value2 = DateTime.Now.AddMonths(2).Date;
         }
         //This method implements the TimePeriod conditional formatting type with LastWeek attribute.
         private void AddTimePeriod_4(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I7:K8"]);
             sheet.Range["I7:K8"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I7:K8"].Style.Color = Color.Linen;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.LastWeek);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I7"];
             c.Value2 = DateTime.Now.AddDays(-6).Date;

             c = sheet.Range["J7"];
             c.Value2 = DateTime.Now.AddDays(-5).Date;

             c = sheet.Range["K7"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I8"];
             c.Text = "LastWeek";

             c = sheet.Range["J8"];
             c.Value2 = DateTime.Now.AddDays(3).Date;

             c = sheet.Range["K8"];
             c.Value2 = DateTime.Now.AddMonths(4).Date;
         }
         //This method implements the TimePeriod conditional formatting type with LastMonth attribute.
         private void AddTimePeriod_3(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I5:K6"]);
             sheet.Range["I5:K6"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I5:K6"].Style.Color = Color.Linen;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.LastMonth);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I5"];
             c.Value2 = DateTime.Now.AddDays(-6).Date;

             c = sheet.Range["J5"];
             c.Value2 = DateTime.Now.AddMonths(-1).Date;

             c = sheet.Range["K5"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I6"];
             c.Text = "LastMonth";

             c = sheet.Range["J6"];
             c.Value2 = DateTime.Now.AddDays(3).Date;

             c = sheet.Range["K6"];
             c.Value2 = DateTime.Now.AddMonths(1).Date;
         }
         //This method implements the TimePeriod conditional formatting type with Last7Days attribute.
         private void AddTimePeriod_2(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I3:K4"]);
             sheet.Range["I3:K4"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I3:K4"].Style.Color = Color.LightSkyBlue;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.Last7Days);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I3"];
             c.Value2 = DateTime.Now.AddDays(-8).Date;

             c = sheet.Range["J3"];
             c.Value2 = DateTime.Now.AddDays(-7).Date;

             c = sheet.Range["K3"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I4"];
             c.Text = "Last7Days";
             
             c = sheet.Range["J4"];
             c.Value2 = DateTime.Now.AddDays(3).Date;

             c = sheet.Range["K4"];
             c.Value2 = DateTime.Now.AddMonths(2).Date;
         }
         //This method implements the TimePeriod conditional formatting type with Today attribute.
         private void AddTimePeriod_1(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["I1:K2"]);
             sheet.Range["I1:K2"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["I1:K2"].Style.Color = Color.LightSlateGray;
             IConditionalFormat cf = conds.AddTimePeriodCondition(TimePeriodType.Today);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["I1"];
             c.Value2 = DateTime.Now.AddDays(-8).Date;

             c = sheet.Range["J1"];
             c.Value2 = DateTime.Now.AddDays(-7).Date;

             c = sheet.Range["K1"];
             c.Value2 = DateTime.Now.Date;

             c = sheet.Range["I2"];
             c.Text = "Today";

             c = sheet.Range["J2"];
             c.Value2 = DateTime.Now.AddDays(3).Date;

             c = sheet.Range["K2"];
             c.Value2 = DateTime.Now.AddMonths(2).Date;
         }
         //This method implements the DuplicateValues conditional formatting type.
         private void AddDuplicate(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E23:G24"]);
             sheet.Range["E23:G24"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E23:G24"].Style.Color = Color.LightSlateGray;
             IConditionalFormat cf = conds.AddDuplicateValuesCondition();
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;
             CellRange c = sheet.Range["E23"];
             c.Text = "aa";
             c = sheet.Range["F23"];
             c.Text = "bb";
             c = sheet.Range["G23"];
             c.Text = "aa";
             c = sheet.Range["E24"];
             c.Text = "bbb";
             c = sheet.Range["F24"];
             c.Text = "bb";
             c = sheet.Range["G24"];
             c.Text = "ccc";
         }
         //This method implements the UniqueValues conditional formatting type.
         private void AddUnique(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E21:G22"]);
             sheet.Range["E21:G22"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E21:G22"].Style.Color = Color.LightSalmon;
             IConditionalFormat cf = conds.AddUniqueValuesCondition();
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;
             CellRange c = sheet.Range["E21"];
             c.Text = "aa";
             c = sheet.Range["F21"];
             c.Text = "bb";
             c = sheet.Range["G21"];
             c.Text = "aa";
             c = sheet.Range["E22"];
             c.Text = "bbb";
             c = sheet.Range["F22"];
             c.Text = "bb";
             c = sheet.Range["G22"];
             c.Text = "ccc";
         }
         //This method implements the NotContainsError conditional formatting type.
         private void AddNotContainsError(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E19:G20"]);
             sheet.Range["E19:G20"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E19:G20"].Style.Color = Color.LightSeaGreen;
             IConditionalFormat cf = conds.AddNotContainsErrorsCondition();
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;
             CellRange c = sheet.Range["E19"];
             c.Text = "aa";
             c = sheet.Range["F19"];
             c.Text = "=Sum";
             c = sheet.Range["G19"];
             c.Text = "aa";
             c = sheet.Range["E20"];
             c.Text = "bbb";
             c = sheet.Range["F20"];
             c.Text = "sss";
             c = sheet.Range["G20"];
             c.Text = "=Max";
         }
         //This method implements the ContainsErrors conditional formatting type.
         private void AddContainsError(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E17:G18"]);
             sheet.Range["E17:G18"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E17:G18"].Style.Color = Color.LightSkyBlue;
             IConditionalFormat cf = conds.AddContainsErrorsCondition();
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;
             CellRange c = sheet.Range["E17"];
             c.Text = "aa";
             c = sheet.Range["F17"];
             c.Text = "=Sum";
             c = sheet.Range["G17"];
             c.Text = "aa";
             c = sheet.Range["E18"];
             c.Text = "bbb";
             c = sheet.Range["F18"];
             c.Text = "sss";
             c = sheet.Range["G18"];
             c.Text = "=Max";
         }
         //This method implements the BeginWith conditional formatting type.
         private void AddBeginWith(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E15:G16"]);
             sheet.Range["E15:G16"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E15:G16"].Style.Color = Color.LightGoldenrodYellow;
             IConditionalFormat cf = conds.AddBeginsWithCondition("ab");
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;

             CellRange c = sheet.Range["E15"];
             c.Text = "aa";
             c = sheet.Range["F15"];
             c.Text = "abc";
             c = sheet.Range["G15"];
             c.Text = "aa";
             c = sheet.Range["E16"];
             c.Text = "bbb";
             c = sheet.Range["F16"];
             c.Text = "sss";
             c = sheet.Range["G16"];
             c.Text = "abcd";
         }
         //This method implements the EndWith conditional formatting type.
         private void AddEndWith(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E13:G14"]);
             sheet.Range["E13:G14"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E13:G14"].Style.Color = Color.LightGray;
             IConditionalFormat cf = conds.AddEndsWithCondition("ab");
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;

             CellRange c = sheet.Range["E13"];
             c.Text = "aa";
             c = sheet.Range["F13"];
             c.Text = "abc";
             c = sheet.Range["G13"];
             c.Text = "aab";
             c = sheet.Range["E14"];
             c.Text = "bbbc";
             c = sheet.Range["F14"];
             c.Text = "sab";
             c = sheet.Range["G14"];
             c.Text = "abcd";
         }
         //This method implements the NotContainsBlank conditional formatting type.
         private void AddNotContainsBlank(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E11:G12"]);
             sheet.Range["E11:G12"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E11:G12"].Style.Color = Color.LightCoral;
             IConditionalFormat cf = conds.AddNotContainsBlanksCondition();
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;

             CellRange c = sheet.Range["E11"];
             c.Text = "aa";
             c = sheet.Range["F11"];
             c.Text = "  ";
             c = sheet.Range["G11"];
             c.Text = "aab";
             c = sheet.Range["E12"];
             c.Text = "abc";
             c = sheet.Range["F12"];
             c.Text = "  ";
             c = sheet.Range["G12"];
             c.Text = "abcd";
         }
          //This method implements the ContainsBlank conditional formatting type.
         private void AddContainsBlank(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E9:G10"]);
             sheet.Range["E9:G10"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E9:G10"].Style.Color = Color.LightCyan;
             IConditionalFormat cf = conds.AddContainsBlanksCondition();
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;

             CellRange c = sheet.Range["E9"];
             c.Text = "aa";
             c = sheet.Range["F9"];
             c.Text = "  ";
             c = sheet.Range["G9"];
             c.Text = "aab";
             c = sheet.Range["E10"];
             c.Text = "abc";
             c = sheet.Range["F10"];
             c.Text = "dvdf";
             c = sheet.Range["G10"];
             c.Text = "abcd";
         }
        //This method implements the NotContainsText conditional formatting type.
         private void AddNotContainsText(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E7:G8"]);
             sheet.Range["E7:G8"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E7:G8"].Style.Color = Color.LightGreen;
             IConditionalFormat cf = conds.AddNotContainsTextCondition("abc");
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;

             CellRange c = sheet.Range["E7"];
             c.Text = "aa";
             c = sheet.Range["F7"];
             c.Text = "abfd";
             c = sheet.Range["G7"];
             c.Text = "aab";
             c = sheet.Range["E8"];
             c.Text = "abc";
             c = sheet.Range["F8"];
             c.Text = "cedf";
             c = sheet.Range["G8"];
             c.Text = "abcd";
         }
        
         //This method implements the ContainsText conditional formatting type.
         private void AddContainsText(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["E5:G6"]);
             sheet.Range["E5:G6"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E5:G6"].Style.Color = Color.LightBlue;
             IConditionalFormat cf = conds.AddContainsTextCondition("abc");
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;

             CellRange c = sheet.Range["E5"];
             c.Text = "aa";
             c = sheet.Range["F5"];
             c.Text = "abfd";
             c = sheet.Range["G5"];
             c.Text = "aab";
             c = sheet.Range["E6"];
             c.Text = "abc";
             c = sheet.Range["F6"];
             c.Text = "cedf";
             c = sheet.Range["G6"];
             c.Text = "abcd";
         }
         //This method implements the DataBars conditional formatting type with Percentile attribute.
         private void AddDataBar2(Worksheet sheet)
         {
             //Add data bars
             XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
             xcfs.AddRange(sheet.Range["E3:G4"]);
             sheet.Range["E3:G4"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E3:G4"].Style.Color = Color.LightGreen;
             IConditionalFormat cf = xcfs.AddCondition();
             cf.FormatType = ConditionalFormatType.DataBar;
             cf.DataBar.BarColor = Color.Orange;
             cf.DataBar.MinPoint.Type = ConditionValueType.Percentile;
             cf.DataBar.MinPoint.Value = 30.78;
             cf.DataBar.ShowValue = false;

             CellRange c = sheet.Range["E3"];
             c.NumberValue = 6;
             c = sheet.Range["F3"];
             c.NumberValue = 9;
             c = sheet.Range["G3"];
             c.NumberValue = 12;
             c = sheet.Range["E4"];
             c.NumberValue = 8;
             c = sheet.Range["F4"];
             c.NumberValue = 11;
             c = sheet.Range["G4"];
             c.NumberValue = 14;
         }
         //This method implements the DataBars conditional formatting type.
         private void AddDataBar1(Worksheet sheet)
         {
             //Add data bars
             XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
             xcfs.AddRange(sheet.Range["E1:G2"]);
             sheet.Range["E1:G2"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["E1:G2"].Style.Color = Color.YellowGreen;
             IConditionalFormat cf = xcfs.AddCondition();
             cf.FormatType = ConditionalFormatType.DataBar;
             cf.DataBar.BarColor = Color.Blue;
             cf.DataBar.MinPoint.Type = ConditionValueType.Percent;
             cf.DataBar.ShowValue = true;

             CellRange c = sheet.Range["E1"];
             c.NumberValue = 4;
             c = sheet.Range["F1"];
             c.NumberValue = 7;
             c = sheet.Range["G1"];
             c.NumberValue = 10;
             c = sheet.Range["E2"];
             c.NumberValue = 6;
             c = sheet.Range["F2"];
             c.NumberValue = 9;
             c = sheet.Range["G2"];
             c.NumberValue = 14;
         }
         //This method implements the IconSet conditional formatting type.
         private void AddDefaultIconSet(Worksheet sheet)
         {
             XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
             xcfs.AddRange(sheet.Range["A1:C2"]);
             sheet.Range["A1:C2"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A1:C2"].Style.Color = Color.Yellow;
             IConditionalFormat cf = xcfs.AddCondition();
             cf.FormatType = ConditionalFormatType.IconSet;
             sheet.Range["A1"].NumberValue = 0;
             sheet.Range["B1"].NumberValue = 3;
             sheet.Range["C1"].NumberValue = 6;
             sheet.Range["A2"].NumberValue = 2;
             sheet.Range["B2"].NumberValue = 5;
             sheet.Range["C2"].NumberValue = 8;
         }
         //This method implements the ColorScale conditional formatting type.
         private void AddDefaultColorScale(Worksheet sheet)
         {
             XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
             xcfs.AddRange(sheet.Range["A5:C6"]);
             sheet.Range["A5:C6"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A5:C6"].Style.Color = Color.Pink;
             IConditionalFormat cf = xcfs.AddCondition();
             cf.FormatType = ConditionalFormatType.ColorScale;

             sheet.Range["A5"].NumberValue = 4;
             sheet.Range["B5"].NumberValue = 7;
             sheet.Range["C5"].NumberValue = 10;
             sheet.Range["A6"].NumberValue = 6;
             sheet.Range["B6"].NumberValue = 9;
             sheet.Range["C6"].NumberValue = 12;
         }
         //This method implements the ColorScale conditional formatting type with some color scale attributes.
         private void Add3ColorScale(Worksheet sheet)
         {
             XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
             xcfs.AddRange(sheet.Range["A7:C8"]);
             sheet.Range["A7:C8"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A7:C8"].Style.Color = Color.Green;
             IConditionalFormat cf = xcfs.AddCondition();
             cf.FormatType = ConditionalFormatType.ColorScale;
             cf.ColorScale.MinValue.Type = ConditionValueType.Number;
             cf.ColorScale.MinValue.Value = 9;
             cf.ColorScale.MinColor = Color.Purple;

             sheet.Range["A7"].NumberValue = 6;
             sheet.Range["B7"].NumberValue = 9;
             sheet.Range["C7"].NumberValue = 12;
             sheet.Range["A8"].NumberValue = 8;
             sheet.Range["B8"].NumberValue = 11;
             sheet.Range["C8"].NumberValue = 14;
         }
         //This method implements the ColorScale conditional formatting type with some color scale attributes.
         private void Add2ColorScale(Worksheet sheet)
         {
             XlsConditionalFormats xcfs = sheet.ConditionalFormats.Add();
             xcfs.AddRange(sheet.Range["A9:C10"]);
             sheet.Range["A9:C10"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A9:C10"].Style.Color = Color.White;
             IConditionalFormat cf = xcfs.AddCondition();
             cf.FormatType = ConditionalFormatType.ColorScale;
             cf.ColorScale.MinColor = Color.Gold;
             cf.ColorScale.MaxColor = Color.SkyBlue;

             sheet.Range["A9"].NumberValue = 8;
             sheet.Range["B9"].NumberValue = 12;
             sheet.Range["C9"].NumberValue = 13;
             sheet.Range["A10"].NumberValue = 10;
             sheet.Range["B10"].NumberValue = 13;
             sheet.Range["C10"].NumberValue = 16;
         }
         //This method implements the AboveAverage conditional formatting type.
         private void AddAboveAverage(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A11:C12"]);
             sheet.Range["A11:C12"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A11:C12"].Style.Color = Color.Tomato;
             IConditionalFormat cf = conds.AddAverageCondition(AverageType.Above);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;

             sheet.Range["A11"].NumberValue = 10;
             sheet.Range["B11"].NumberValue = 13;
             sheet.Range["C11"].NumberValue = 16;
             sheet.Range["A12"].NumberValue = 12;
             sheet.Range["B12"].NumberValue = 15;
             sheet.Range["C12"].NumberValue = 18;
         }
         //This method implements an BelowEqualAverage conditional formatting type with some custom attributes.
         private void AddAboveAverage2(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A13:C14"]);
             sheet.Range["A13:C14"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A13:C14"].Style.Color = Color.LightPink;
             IConditionalFormat cf = conds.AddAverageCondition(AverageType.BelowEqual);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.LightSkyBlue;

             sheet.Range["A13"].NumberValue = 12;
             sheet.Range["B13"].NumberValue = 15;
             sheet.Range["C13"].NumberValue = 18;
             sheet.Range["A14"].NumberValue = 14;
             sheet.Range["B14"].NumberValue = 17;
             sheet.Range["C14"].NumberValue = 20;
         }
         // This method implements an AboveStdDev3 conditional formatting type with some custom attributes.
         private void AddAboveAverage3(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A15:C16"]);
             sheet.Range["A15:C16"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A15:C16"].Style.Color = Color.LightPink;
             IConditionalFormat cf = conds.AddAverageCondition(AverageType.AboveStdDev3);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.LightSkyBlue;
            
             sheet.Range["A15"].NumberValue = 12;
             sheet.Range["B15"].NumberValue = 15;
             sheet.Range["C15"].NumberValue = 18;
             sheet.Range["A16"].NumberValue = 14;
             sheet.Range["B16"].NumberValue = 17;
             sheet.Range["C16"].NumberValue = 20;
         }
         //This method implements a Top10 conditional formatting type.
         private void AddTop10_1(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A17:C20"]);
             sheet.Range["A17:C20"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A17:C20"].Style.Color = Color.Gray;
             IConditionalFormat cf = conds.AddTopBottomCondition(TopBottomType.Top, 10);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Yellow;

             sheet.Range["A17"].NumberValue = 16;
             sheet.Range["B17"].NumberValue = 21;
             sheet.Range["C17"].NumberValue = 26;
             sheet.Range["A18"].NumberValue = 18;
             sheet.Range["B18"].NumberValue = 23;
             sheet.Range["C18"].NumberValue = 28;
             sheet.Range["A19"].NumberValue = 20;
             sheet.Range["B19"].NumberValue = 25;
             sheet.Range["C19"].NumberValue = 30;
             sheet.Range["A20"].NumberValue = 22;
             sheet.Range["B20"].NumberValue = 27;
             sheet.Range["C20"].NumberValue = 32;
         }
         //This method implements Bottom 10 conditional formatting type.
         private void AddTop10_2(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A21:C24"]);
             sheet.Range["A21:C24"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A21:C24"].Style.Color = Color.Green;
             IConditionalFormat cf = conds.AddTopBottomCondition(TopBottomType.Bottom, 10);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Pink;

             sheet.Range["A21"].NumberValue = 20;
             sheet.Range["B21"].NumberValue = 25;
             sheet.Range["C21"].NumberValue = 30;
             sheet.Range["A22"].NumberValue = 22;
             sheet.Range["B22"].NumberValue = 27;
             sheet.Range["C22"].NumberValue = 32;
             sheet.Range["A23"].NumberValue = 24;
             sheet.Range["B23"].NumberValue = 29;
             sheet.Range["C23"].NumberValue = 34;
             sheet.Range["A24"].NumberValue = 24;
             sheet.Range["B24"].NumberValue = 31;
             sheet.Range["C24"].NumberValue = 36;
         }
         //This method implements TopPercent 10 conditional formatting type with some custom attributes.
         private void AddTop10_3(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A25:C28"]);
             sheet.Range["A25:C28"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A25:C28"].Style.Color = Color.Orange;
             IConditionalFormat cf = conds.AddTopBottomCondition(TopBottomType.TopPercent, 10);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Blue;

             sheet.Range["A25"].NumberValue = 24;
             sheet.Range["B25"].NumberValue = 29;
             sheet.Range["C25"].NumberValue = 34;
             sheet.Range["A26"].NumberValue = 25;
             sheet.Range["B26"].NumberValue = 36;
             sheet.Range["C26"].NumberValue = 32;
             sheet.Range["A27"].NumberValue = 24;
             sheet.Range["B27"].NumberValue = 28;
             sheet.Range["C27"].NumberValue = 31;
             sheet.Range["A28"].NumberValue = 34;
             sheet.Range["B28"].NumberValue = 26;
             sheet.Range["C28"].NumberValue = 32;
         }
         //This method implements BottomPercent 10 conditional formatting type with some custom attributes.
         private void AddTop10_4(Worksheet sheet)
         {
             XlsConditionalFormats conds = sheet.ConditionalFormats.Add();
             conds.AddRange(sheet.Range["A29:C32"]);
             sheet.Range["A29:C32"].Style.FillPattern = ExcelPatternType.Solid;
             sheet.Range["A29:C32"].Style.Color = Color.Gold;
             IConditionalFormat cf = conds.AddTopBottomCondition(TopBottomType.BottomPercent, 10);
             cf.FillPattern = ExcelPatternType.Solid;
             cf.BackColor = Color.Green;

             sheet.Range["A29"].NumberValue = 22;
             sheet.Range["B29"].NumberValue = 33;
             sheet.Range["C29"].NumberValue = 38;
             sheet.Range["A30"].NumberValue = 30;
             sheet.Range["B30"].NumberValue = 35;
             sheet.Range["C30"].NumberValue = 39;
             sheet.Range["A31"].NumberValue = 32;
             sheet.Range["B31"].NumberValue = 37;
             sheet.Range["C31"].NumberValue = 43;
             sheet.Range["A32"].NumberValue = 34;
             sheet.Range["B32"].NumberValue = 28;
             sheet.Range["C32"].NumberValue = 32;
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
