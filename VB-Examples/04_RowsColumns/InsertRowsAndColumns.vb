Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace InsertRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\InsertRowsAndColumns.xls")

			Dim worksheet As Worksheet = workbook.Worksheets(0)
			'Inserting a row into the worksheet 
			worksheet.InsertRow(2)
			'Inserting a column into the worksheet 
			worksheet.InsertColumn(2)
			'Inserting multiple rows into the worksheet
			worksheet.InsertRow(5, 2)
			'Inserting multiple columns into the worksheet
			worksheet.InsertColumn(5, 2)

			Dim result As String="InsertRowsAndColumns_out.xlsx"
			workbook.SaveToFile(result,ExcelVersion.Version2010)
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
