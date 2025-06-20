Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace GetCellDataType
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_2.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Iterate over each cell range in the specified range.
            For Each range As CellRange In sheet.Range("H2:H7")
                'Retrieve the type of the cell at the specified row and column.
                Dim cellType As XlsWorksheet.TRangeValueType = sheet.GetCellType(range.Row, range.Column, False)
                'Set the text value of the adjacent cell in the same row to the cell type.
                sheet(range.Row, range.Column + 1).Text = cellType.ToString()
                'Set the font color of the adjacent cell in the same row to red.
                sheet(range.Row, range.Column + 1).Style.Font.Color = Color.Red
                'Apply bold formatting to the font of the adjacent cell in the same row.
                sheet(range.Row, range.Column + 1).Style.Font.IsBold = True
            Next range
            'Specify the output file name.
            Dim result As String = "Result-GetCellDataType.xlsx"

            'Save the modified workbook to a file in Excel 2013 format.
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
