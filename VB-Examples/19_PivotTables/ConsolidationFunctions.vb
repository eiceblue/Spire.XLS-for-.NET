Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ConsolidationFunctions
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load an excel file including pivot table
			workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")
			'Get the sheet in which the pivot table is located
			Dim sheet As Worksheet = workbook.Worksheets("PivotTable")

			Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

			'Apply Average consolidation function to first data field
			pt.DataFields(0).Subtotal = SubtotalTypes.Average

			'Apply Max consolidation function to second data field
			pt.DataFields(1).Subtotal = SubtotalTypes.Max

			pt.CalculateData()

			Dim result As String = "ConsolidationFunctions_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
