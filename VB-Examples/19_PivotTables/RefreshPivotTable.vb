Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace RefreshPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook.
			Dim workbook As New Workbook()

			'Load the file from disk.
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_7.xlsx")

			'Get the first worksheet.
			Dim sheet As Worksheet = workbook.Worksheets(1)

			'Update the data source of PivotTable.
			sheet.Range("D2").Value = "999"

			'Get the PivotTable that was built on the data source.
			Dim pt As XlsPivotTable = TryCast(workbook.Worksheets(0).PivotTables(0), XlsPivotTable)

			'Refresh the data of PivotTable.
			pt.Cache.IsRefreshOnLoad = True

			Dim result As String = "Result-RefreshPivotTable.xlsx"

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
