Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace RepeatAllItemLabelsForPivotTable
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new instance of the Workbook class
			Dim workbook As New Workbook()

			' Load the workbook from the specified file path 
			workbook.LoadFromFile("..\..\..\..\..\..\Data\RepeatAllItemLabelsForPivotTable.xlsx")

			' Iterate through each pivot table in the "Pivot" worksheet
			For Each pt As XlsPivotTable In workbook.Worksheets("Pivot").PivotTables
				' Set the RepeatAllItemLabels property to true for the pivot table
				pt.Options.RepeatAllItemLabels = True

				' Calculate the data for the pivot table
				pt.CalculateData()

				' Refresh the cache for the pivot table
				pt.Cache.IsRefreshOnLoad = True
			Next pt

			' Define the output file name for the modified workbook
			Dim result As String = "RepeatAllItemLabelsForPivotTable_output.xlsx"

			' Save the modified workbook to the specified file path with the specified Excel version
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose the workbook instance to release resources
			workbook.Dispose()

			' Launch the file
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
