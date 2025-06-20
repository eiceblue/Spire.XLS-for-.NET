Imports Spire.Xls

Namespace ChangeDataSource
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load an Excel file from a specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChangeDataSource.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Define a CellRange object that represents the range "A1:C15" on the worksheet.
            Dim Range As CellRange = sheet.Range("A1:C15")

            ' Get the first PivotTable from the second worksheet in the workbook and cast it to a PivotTable object.
            Dim table As PivotTable = TryCast(workbook.Worksheets(1).PivotTables(0), PivotTable)

            ' Change the data source of the PivotTable to the specified range.
            table.ChangeDataSource(Range)

            ' Disable automatic refresh of the PivotTable cache when the workbook is loaded.
            table.Cache.IsRefreshOnLoad = False

            ' Specify the name for the result file.
            Dim result As String = "ChangeDataSource_result.xlsx"

            ' Save the modified workbook to a file with the specified result name and Excel version.
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
