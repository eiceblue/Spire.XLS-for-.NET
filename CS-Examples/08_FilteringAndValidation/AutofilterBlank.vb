Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace AutofilterBlank
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "AutofilterBlank.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AutofilterBlank.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Enable the auto-filter feature and filter for blank cells in the first column (column index 0)
            sheet.AutoFilters.MatchBlanks(0)

            ' Apply the filter based on the specified criteria
            sheet.AutoFilters.Filter()

            ' Specify the output filename for the filtered data
            Dim output As String = "AutofilterBlank_out.xlsx"

            ' Save the filtered data to a new Excel file with Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
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
