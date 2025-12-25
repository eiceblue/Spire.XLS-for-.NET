Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ExportDataKeepDataType
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a workbook
			Dim workbook As New Workbook()

			' Load the file from disk
			workbook.LoadFromFile("../../../../../../Data/ExportDataKeepDataType.xlsx")

			' Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Export DataTable without keeping data type
			Dim options As New ExportTableOptions()
			options.ExportColumnNames = True
			options.KeepDataFormat = False
			options.KeepDataType = True
			options.RenameStrategy = RenameStrategy.Digit

			' Export data to data table
			Dim table As DataTable = sheet.ExportDataTable(1, 1, sheet.LastDataRow, sheet.LastDataColumn, options)

			' Show the data table
			Me.dataGridView1.DataSource = table

			' Dispose of the workbook object to free up resources
			workbook.Dispose()
		End Sub
		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
