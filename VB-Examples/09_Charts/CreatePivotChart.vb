Imports Spire.Xls
Imports Spire.Xls.Core

Namespace CreatePivotChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTable.xlsx")

			'get the first worksheet
		   Dim sheet As Worksheet = workbook.Worksheets(0)
		   'get the first pivot table in the worksheet
		   Dim pivotTable As IPivotTable = sheet.PivotTables(0)

		   'create a clustered column chart based on the pivot table
		   Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable)
		   'set chart position
		   chart.TopRow = 12
		   chart.LeftColumn = 1
		   chart.RightColumn = 8
		   chart.BottomRow = 30
		   chart.ChartTitle = "Product"
		   chart.PrimaryCategoryAxis.MultiLevelLable = True

			'Save the document
			Dim output As String = "CreatePivotChart.xlsx"
		workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the file
			ExcelDocViewer(output)
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
