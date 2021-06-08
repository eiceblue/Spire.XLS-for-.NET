Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace HideRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\HideRowsAndColumns.xls")

			Dim worksheet As Worksheet = workbook.Worksheets(0)
			' Hiding the column of the worksheet
			worksheet.HideColumn(2)
			'Hiding the row of the worksheet
			worksheet.HideRow(4)

			workbook.SaveToFile("HideRowsAndColumns.xlsx", ExcelVersion.Version2010)

			ExcelDocViewer("HideRowsAndColumns.xlsx")
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
