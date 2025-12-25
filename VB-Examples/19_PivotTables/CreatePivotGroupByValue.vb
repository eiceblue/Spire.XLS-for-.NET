Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace CreatePivotGroupByValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new workbook object
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\CreatePivotGroupByValue.xlsx")

			' Get the reference to the first sheet in the workbook
			Dim pivotSheet As Worksheet = workbook.Worksheets(0)

			' Cast the first PivotTable in the PivotTables collection to an XlsPivotTable object.
			Dim pivot As XlsPivotTable = CType(pivotSheet.PivotTables(0), XlsPivotTable)

			' Retrieve the PivotField named "number" from the PivotTable and cast it to a PivotField object.
			Dim dateBaseField As PivotField = TryCast(pivot.PivotFields("number"), PivotField)

			' Create a group for the PivotField, starting at 3000, ending at 3800, with an interval of 1.
			dateBaseField.CreateGroup(3000, 3800, 1)

			' Recalculate the data in the PivotTable to reflect the changes made.
			pivot.CalculateData()

			' Specify the filename for the resulting Excel file
			Dim result As String = "CreatePivotGroupByValue-out.xlsx"

			' Save the workbook to the specified file in Excel 2010 format
			workbook.SaveToFile(result, ExcelVersion.Version2016)

			' Dispose of the workbook object
			workbook.Dispose()

			' View the document using a file viewer
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
