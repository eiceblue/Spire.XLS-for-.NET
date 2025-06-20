Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports Spire.Xls.Core

Namespace GetAndClearCellCurrentRegion
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel file from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_10.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Retrieve the current region of cells starting from cell A1.
            Dim xlRange As IXLSRange = sheet.Range("A1").CurrentRegion
            'Iterate through each cell range in the current region.
            For Each range As CellRange In xlRange
                'Clear all contents and formatting in the current cell range.
                range.ClearAll()
            Next range
            'Specify the file name for the output file.
            Dim result As String = "CellCurrentRegion_result.xlsx"

            'Save the modified workbook to the specified output file using Excel 2016 version.
            workbook.SaveToFile(result, ExcelVersion.Version2016)
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
