Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace ExpandOrCollapseRows
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
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the data in Pivot Table.
			Dim pivotTable As Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable = TryCast(sheet.PivotTables(0), Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotTable)

			'Calculate Data.
			pivotTable.CalculateData()

			'Collapse the rows.
			TryCast(pivotTable.PivotFields("Vendor No"), Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3501", True)

			'Expand the rows.
			TryCast(pivotTable.PivotFields("Vendor No"), Spire.Xls.Core.Spreadsheet.PivotTables.XlsPivotField).HideItemDetail("3502", False)

			Dim result As String = "Result-ExpandOrCollapseRowsInPivotTable.xlsx"

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
