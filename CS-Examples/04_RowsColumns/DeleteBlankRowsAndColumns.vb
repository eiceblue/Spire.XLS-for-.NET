Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace DeleteBlankRowsAndColumns
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_2.xlsx")

            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Iterates over the rows of the worksheet in reverse order.
            For i As Integer = sheet.Rows.Length - 1 To 0 Step -1
                'Checks if the current row is blank.
                If sheet.Rows(i).IsBlank Then
                    'Deletes the current row from the worksheet.
                    sheet.DeleteRow(i + 1)
                End If
            Next i

            'Iterates over the columns of the worksheet in reverse order.
            For j As Integer = sheet.Columns.Length - 1 To 0 Step -1
                'Checks if the current column is blank.
                If sheet.Columns(j).IsBlank Then
                    'Deletes the current column from the worksheet.
                    sheet.DeleteColumn(j + 1)
                End If
            Next j

            'Specifies the name of the resulting Excel file.
            Dim result As String = "Result-DeleteBlankRowsAndColumns.xlsx"

            'Saves the modified workbook to a file with the specified name and Excel version.
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
