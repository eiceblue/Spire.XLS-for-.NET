Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace OnlyCopyFormulaValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CopyOnlyFormulaValue1.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set the copy option
			Dim copyOptions As CopyRangeOptions = CopyRangeOptions.OnlyCopyFormulaValue

			Dim sourceRange As CellRange = sheet.Range("A6:E6")
			sheet.Copy(sourceRange, sheet.Range("A8:E8"), copyOptions)


			sourceRange.Copy(sheet.Range("A10:E10"), copyOptions)

			Dim result As String = "Result-OnlyCopyFormulaValue.xlsx"

			'Save to file.
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the MS Excel file.
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
