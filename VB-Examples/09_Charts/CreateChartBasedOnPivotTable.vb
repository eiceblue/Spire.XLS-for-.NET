Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.ComponentModel
Imports System.Text

Namespace CreateChartBasedOnPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load an excel file including pivot table
			workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTable.xlsx")

			'Get the sheet in which the pivot table is located
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim pt As XlsPivotTable = TryCast(sheet.PivotTables(0), XlsPivotTable)

			workbook.Worksheets(1).Charts.Add(ExcelChartType.BarClustered, pt)

			'Save the document
			Dim output As String = "CreateChartBasedOnPivotTable.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'View the document
			FileViewer(output)
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
