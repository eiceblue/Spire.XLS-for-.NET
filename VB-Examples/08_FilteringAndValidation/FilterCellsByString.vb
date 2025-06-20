Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.AutoFilter

Namespace FilterCellsByString
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "Template_Xls_6.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_6.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the range for auto-filtering to be column D (cells D1 to D19) in the worksheet
            sheet.AutoFilters.Range = sheet.Range("D1:D19")

            ' Retrieve the filter column for column D
            Dim filtercolumn As FilterColumn = CType(sheet.AutoFilters(0), FilterColumn)

            ' Apply a custom filter to filter cells in column D that start with "South"
            sheet.AutoFilters.CustomFilter(filtercolumn, FilterOperatorType.Equal, "South*")

            ' Apply the filter based on the specified criteria
            sheet.AutoFilters.Filter()

            ' Save the filtered data to a new Excel file named "filterCellsByString_result.xlsx" with Excel 2013 format
            workbook.SaveToFile("filterCellsByString_result.xlsx", ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            FileViewer("filterCellsByString_result.xlsx")
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
