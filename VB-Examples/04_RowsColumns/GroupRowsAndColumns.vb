Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace GroupRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\GroupRowsAndColumns.xls")
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Grouping rows
			sheet.GroupByRows(1,5,False)
			'Grouping columns
			sheet.GroupByColumns(1,3,False)

			workbook.SaveToFile("GroupRowsAndColumns.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("GroupRowsAndColumns.xlsx")
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
