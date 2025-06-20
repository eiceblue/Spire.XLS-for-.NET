Imports System.IO

Imports Spire.Xls

Namespace CalculateFormulas
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

        Private Function WriteFormulas() As Stream
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Initialize variables for row and formula tracking
            Dim currentRow As Integer = 1
            Dim currentFormula As String = String.Empty

            ' Set column widths for specific columns in the sheet
            sheet.SetColumnWidth(1, 32)
            sheet.SetColumnWidth(2, 16)
            sheet.SetColumnWidth(3, 16)

            ' Set value for a cell in the sheet
            sheet.Range(currentRow, 1).Value = "Examples of formulas :"
            currentRow += 1
            currentRow += 1
            sheet.Range(currentRow, 1).Value = "Test data:"

            ' Apply styling to a range of cells
            Dim range As CellRange = sheet.Range("A1")
            range.Style.Font.IsBold = True
            range.Style.FillPattern = ExcelPatternType.Solid
            range.Style.KnownColor = ExcelColors.LightGreen1
            range.Style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium


            ' Set numeric values for cells in the sheet
            sheet.Range(currentRow, 2).NumberValue = 7.3
            sheet.Range(currentRow, 3).NumberValue = 5

            sheet.Range(currentRow, 4).NumberValue = 8.2
            sheet.Range(currentRow, 5).NumberValue = 4
            sheet.Range(currentRow, 6).NumberValue = 3
            sheet.Range(currentRow, 7).NumberValue = 11.3

            currentRow += 1
            sheet.Range(currentRow, 1).Value = "Formulas"
            sheet.Range(currentRow, 2).Value = "Results"
            range = sheet.Range(currentRow, 1, currentRow, 2)
            range.Style.Font.IsBold = True
            range.Style.KnownColor = ExcelColors.LightGreen1
            range.Style.FillPattern = ExcelPatternType.Solid
            range.Style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium

            ' Set string values for cells in the sheet
            currentFormula = "=""hello"""
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            sheet.Range(currentRow, 3).Formula = "=""" & New String(New Char() {ChrW(&H4F60), ChrW(&H597D)}) & """"

            ' Set int values for cells in the sheet
            currentFormula = "=300"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            ' Set float values for cells in the sheet
            currentFormula = "=3389.639421"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            ' Set bool values for cells in the sheet
            currentFormula = "=false"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            currentFormula = "=1+2+3+4+5-6-7+8-9"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            currentFormula = "=33*3/4-2+10"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula


            currentFormula = "=Sheet1!$B$3"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            currentFormula = "=Count(3,5,8,10,2,34)"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula

            currentFormula = "=NOW()"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            sheet.Range(currentRow, 2).Style.NumberFormat = "yyyy-MM-DD"

            currentFormula = "=SECOND(11)"
            currentRow += 1
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=MINUTE(12)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=MONTH(9)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=DAY(10)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=TIME(4,5,7)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=DATE(6,4,2)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=RAND()"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=HOUR(12)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=MOD(5,3)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=WEEKDAY(3)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=YEAR(23)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=NOT(true)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=OR(true)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=AND(TRUE)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=VALUE(30)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=LEN(""world"")"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=MID(""world"",4,2)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=ROUND(7,3)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=SIGN(4)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=INT(200)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=ABS(-1.21)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=LN(15)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=EXP(20)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=SQRT(40)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=PI()"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=COS(9)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=SIN(45)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=MAX(10,30)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=MIN(5,7)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=AVERAGE(12,45)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=SUM(18,29)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            currentRow += 1

            currentFormula = "=IF(4,2,2)"
            sheet.Range(currentRow, 1).Text = "'" + currentFormula
            sheet.Range(currentRow, 2).Formula = currentFormula
            workbook.CalculateAllValue()

            Dim buffer As New MemoryStream()
            workbook.SaveToStream(buffer)
            buffer.Position = 0

            Return buffer
        End Function


        Private Sub btnExport_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnExport.Click
			Using buffer As Stream = Me.WriteFormulas()
				'load
				Dim workbook As New Workbook()
				workbook.LoadFromStream(buffer)

				'calculate all cells
				workbook.CalculateAllValue()

				'// calculate the first worksheet
				'workbook.Worksheets[0].CalculateAllValue();

				'// calculate formula of cell A10
				'workbook.Worksheets[0].Range["A10"].CalculateAllValue();

				'// when convert to pdf, please set IsReCalculateOnConvert as false if you use sheet or cell area calculation
				'workbook.ConverterSetting.IsReCalculateOnConvert = false;

				'export
				Dim sheet As Worksheet = workbook.Worksheets(0)
				Me.dataGrid1.DataSource = sheet.ExportDataTable(sheet("A4:B46"), True, True)
				workbook.SaveToFile("result.xlsx",ExcelVersion.Version2013)
			End Using
			ExcelDocViewer("result.xlsx")
		End Sub


		Private Sub btnCalculate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCalculate.Click
			Using buffer As Stream = Me.WriteFormulas()
				'load
				Dim workbook As New Workbook()
				workbook.LoadFromStream(buffer)

				'calculate formula
				Dim b3 As Object = workbook.CalculateFormulaValue("Sheet1!$B$3")
				Dim c3 As Object = workbook.CalculateFormulaValue("Sheet1!$C$3")
				Dim formula As String = "Sheet1!$B$3 + Sheet1!$C$3"
				Dim value As Object = workbook.CalculateFormulaValue(formula)
				Dim message As String = String.Format("Sheet1!$B$3 = {0}, Sheet1!$C$3 = {1}, {2} = {3}", b3, c3, formula, value)
				MessageBox.Show(message)
			End Using
		End Sub

		Private Sub btnRead_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRead.Click
			Using buffer As Stream = Me.WriteFormulas()
				'load
				Dim workbook As New Workbook()
				workbook.LoadFromStream(buffer)

				'calculate all cells' formulas
				workbook.CalculateAllValue()

				'read cells' value to data table
				Dim sheet As Worksheet = workbook.Worksheets(0)
				Dim dataTable As New DataTable()
				dataTable.Columns.Add("Formulas", GetType(String))
				dataTable.Columns.Add("Results", GetType(Object))
				For Each row As CellRange In sheet("A5:B46").Rows
					Dim formula As String = row.Columns(1).Formula
					Dim value As Object = row.Columns(1).FormulaValue
					dataTable.Rows.Add(formula, value)
				Next row
				Me.dataGrid1.DataSource = dataTable
			End Using
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
