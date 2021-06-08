Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace DeleteBlankRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_2.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Delete blank rows from the worksheet.
			For i As Integer = sheet.Rows.Length - 1 To 0 Step -1
				If sheet.Rows(i).IsBlank Then
					sheet.DeleteRow(i + 1)
				End If
			Next i

			'Delete blank columns from the worksheet.
			For j As Integer = sheet.Columns.Length - 1 To 0 Step -1
				If sheet.Columns(j).IsBlank Then
					sheet.DeleteColumn(j + 1)
				End If
			Next j

			Dim result As String = "Result-DeleteBlankRowsAndColumns.xlsx"

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
