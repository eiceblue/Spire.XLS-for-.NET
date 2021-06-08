Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetDataValidationOnSeparateSheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SetDataValidationOnSeparateSheet.xlsx")

			'This is the first sheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			sheet1.Range("B10").Text = "Here is a dataValidation example."

			'This is the second sheet
			Dim sheet2 As Worksheet = workbook.Worksheets(1)

			'The property is to enable the data can be from different sheet.
			sheet2.ParentWorkbook.Allow3DRangesInDataValidation = True
			sheet1.Range("B11").DataValidation.DataRange = sheet2.Range("A1:A7")

			workbook.SaveToFile("result.xlsx",ExcelVersion.Version2013)
			ExcelDocViewer("result.xlsx")
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
