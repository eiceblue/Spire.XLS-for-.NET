Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace ShowDataFieldInRow
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "PivotTableExample.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTableExample.xlsx")

            ' Get the first PivotTable from the first worksheet in the workbook
            Dim pivotTable As XlsPivotTable = TryCast(workbook.Worksheets(1).PivotTables(0), XlsPivotTable)

            ' Set the option to show the data field in the row of the PivotTable
            pivotTable.ShowDataFieldInRow = True

            ' Calculate the data in the PivotTable
            pivotTable.CalculateData()

            ' Save the modified workbook to a new file named "result.xlsx" in Excel 2016 format
            workbook.SaveToFile("result.xlsx", ExcelVersion.Version2016)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("result.xlsx")
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
