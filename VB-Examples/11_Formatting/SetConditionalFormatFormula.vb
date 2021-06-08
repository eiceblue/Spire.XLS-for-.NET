Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Drawing.Imaging
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.Collections

Namespace SetConditionalFormatFormula
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the default first  worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add ConditionalFormat
			Dim xcfs As XlsConditionalFormats = sheet.ConditionalFormats.Add()

			'Define the range
			xcfs.AddRange(sheet.Range("B5"))

			'Add condition
			Dim format As IConditionalFormat = xcfs.AddCondition()
			format.FormatType = ConditionalFormatType.CellValue

			'If greater than 1000
			format.FirstFormula = "1000"
			format.Operator = ComparisonOperatorType.Greater
			format.BackColor = Color.Orange

       sheet.Range("B1").NumberValue=40
       sheet.Range("B2").NumberValue=500
       sheet.Range("B3").NumberValue=300
       sheet.Range("B4").NumberValue=400
			'Set a SUM formula for B5 
			sheet.Range("B5").Formula = "=SUM(B1:B4)"

			'Add text
			sheet.Range("C5").Text = "If Sum of B1:B4 is greater than 1000, B5 will have orange background."

			'String for output file 
			Dim outputFile As String = "Output.xlsx"

			'Save the file
			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
