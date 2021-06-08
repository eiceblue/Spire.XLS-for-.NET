Imports Spire.Xls

Namespace ChangeDataSource
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()
			'Load an excel file
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ChangeDataSource.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim Range As CellRange = sheet.Range("A1:C15")

			Dim table As PivotTable = TryCast(workbook.Worksheets(1).PivotTables(0), PivotTable)

			'Change data source
			table.ChangeDataSource(Range)
			table.Cache.IsRefreshOnLoad = False

			Dim result As String = "ChangeDataSource_result.xlsx"
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
