Imports Spire.Xls
Imports Spire.Xls.Core

Namespace CreatePivotChart

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "PivotTable.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTable.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first PivotTable from the worksheet
            Dim pivotTable As IPivotTable = sheet.PivotTables(0)

            ' Add a column clustered chart to the worksheet, using the PivotTable as the data source
            Dim chart As Chart = sheet.Charts.Add(ExcelChartType.ColumnClustered, pivotTable)

            ' Set the position of the chart within the worksheet using row and column indices
            chart.TopRow = 12
            chart.LeftColumn = 1
            chart.RightColumn = 8
            chart.BottomRow = 30

            ' Set the chart title to "Product"
            chart.ChartTitle = "Product"

            ' Enable multi-level category axis labels
            chart.PrimaryCategoryAxis.MultiLevelLable = True

            ' Specify the output file name as "CreatePivotChart.xlsx"
            Dim output As String = "CreatePivotChart.xlsx"

            ' Save the modified workbook to the specified file path, using Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
