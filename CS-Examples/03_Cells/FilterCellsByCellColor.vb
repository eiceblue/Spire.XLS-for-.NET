Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core.Spreadsheet.AutoFilter

Namespace FilterCellsByCellColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel file from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_3.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Create an auto filter on column G (range G1:G19) in the worksheet.
            sheet.AutoFilters.Range = sheet.Range("G1:G19")

            'Retrieve the filter column for column G.
            Dim filtercolumn As FilterColumn = CType(sheet.AutoFilters(0), FilterColumn)

            'Add a color filter to filter the cells in the filter column that have a red fill color.
            sheet.AutoFilters.AddFillColorFilter(filtercolumn, Color.Red)

            'Apply the filters to the worksheet data based on the specified criteria.
            sheet.AutoFilters.Filter()

            'Specify the file name for the output file.
            Dim result As String = "Result-FilterCellsByCellColor.xlsx"

            'Save the modified workbook to the specified output file using Excel 2013 version.
            workbook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
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
