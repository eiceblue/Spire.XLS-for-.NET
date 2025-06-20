Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace WriteFormulas
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
  
			' Create a new workbook
			Dim workbook As New Workbook()

			' Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim currentRow As Integer = 1
			Dim currentFormula As String = String.Empty

			' Set column widths
			sheet.SetColumnWidth(1, 32)
			sheet.SetColumnWidth(2, 16)
			sheet.SetColumnWidth(3, 16)

			' Add header and test data labels
			sheet.Range(currentRow, 1).Value = "Examples of formulas :"
			currentRow += 1
			currentRow += 1
			sheet.Range(currentRow, 1).Value = "Test data:"

			' Apply formatting to the header cell
			Dim range As CellRange = sheet.Range("A1")
			range.Style.Font.IsBold = True
			range.Style.FillPattern = ExcelPatternType.Solid
			range.Style.KnownColor = ExcelColors.LightGreen1
			range.Style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium

			' test data
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
			' range.Value = "Formulas";
			range.Style.Font.IsBold = True
			range.Style.KnownColor = ExcelColors.LightGreen1
			range.Style.FillPattern = ExcelPatternType.Solid
			range.Style.Borders(BordersLineType.EdgeBottom).LineStyle = LineStyleType.Medium
			' str.
			currentFormula = "=""hello"""
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = "=""hello"""
			sheet.Range(currentRow, 2).Formula = currentFormula
			sheet.Range(currentRow, 3).Formula = "=""" & New String(New Char() { ChrW(&H4f60), ChrW(&H597d) }) & """"

			' int.
			currentFormula = "=300"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula

			' float
			currentFormula = "=3389.639421"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula

			' bool.
			currentFormula = "=false"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula

			currentFormula = "=1+2+3+4+5-6-7+8-9"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula

			currentFormula = "=33*3/4-2+10"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula


			' sheet reference
			currentFormula = "=Sheet1!$B$3"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula

			' sheet area reference
			currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula

			' Functions
			currentFormula = "=Count(3,5,8,10,2,34)"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula


			currentFormula = "=NOW()"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			sheet.Range(currentRow, 2).Style.NumberFormat = "yyyy-MM-DD"

			currentFormula = "=SECOND(11)"
			currentRow += 1
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=MINUTE(12)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=MONTH(9)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=DAY(10)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=TIME(4,5,7)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=DATE(6,4,2)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=RAND()"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=HOUR(12)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=MOD(5,3)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=WEEKDAY(3)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=YEAR(23)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=NOT(true)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=OR(true)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=AND(TRUE)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=VALUE(30)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=LEN(""world"")"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=MID(""world"",4,2)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=ROUND(7,3)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=SIGN(4)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=INT(200)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=ABS(-1.21)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=LN(15)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=EXP(20)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=SQRT(40)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=PI()"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=COS(9)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=SIN(45)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=MAX(10,30)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=MIN(5,7)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=AVERAGE(12,45)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=SUM(18,29)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=IF(4,2,2)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			currentFormula = "=SUBTOTAL(3,Sheet1!B2:E3)"
			sheet.Range(currentRow, 1).NumberFormat = "@"
			sheet.Range(currentRow, 1).Text = currentFormula
			sheet.Range(currentRow, 2).Formula = currentFormula
			currentRow += 1

			' Specify the file name for the resulting Excel file
			Dim result As String = "WriteFormulas_result.xlsx"

			' Save the workbook to the specified file in Excel 2013 format
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
			ExcelDocViewer(result)



		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
