Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace RemoveDataValidation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\RemoveDataValidation.xlsx")

			'Create an array of rectangles, which is used to locate the ranges in worksheet.
			Dim rectangles(0) As Rectangle

			'Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
			rectangles(0) = New Rectangle(0, 0, 1, 2)

			'Remove validations in the ranges represented by rectangles.
			workbook.Worksheets(0).DVTable.Remove(rectangles)

			Dim result As String = "Result-RemoveDataValidation.xlsx"

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
