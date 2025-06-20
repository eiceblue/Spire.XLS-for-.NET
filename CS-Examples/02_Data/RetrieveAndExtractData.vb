Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace RetrieveAndExtractData
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Creates a new Excel workbook.
            Dim newBook As New Workbook()

            ' Retrieves the first worksheet in the new workbook (index starts at 0).
            Dim newSheet As Worksheet = newBook.Worksheets(0)

            ' Creates another new Excel workbook.
            Dim workbook As New Workbook()

            ' Loads an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_3.xlsx")

            ' Retrieves the first worksheet in the loaded workbook (index starts at 0).
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Initializes a counter variable.
            Dim i As Integer = 1
            ' Retrieves the number of columns in the sheet.
            Dim columnCount As Integer = sheet.Columns.Length
            ' Iterates through each cell in the first column of the sheet.
            For Each range As CellRange In sheet.Columns(0)
                ' Checks if the cell text is "teacher".
                If range.Text = "teacher" Then
                    ' Defines the source range to copy, which spans from the current row to all columns.
                    Dim sourceRange As CellRange = sheet.Range(range.Row, 1, range.Row, columnCount)
                    ' Defines the destination range where the data will be copied to in the new workbook.
                    Dim destRange As CellRange = newSheet.Range(i, 1, i, columnCount)
                    ' Copies the data from the source range to the destination range in the new workbook, including formatting and formulas.
                    sheet.Copy(sourceRange, destRange, True)
                    ' Increments the counter variable.
                    i += 1
                End If
            Next range
            ' Defines the file name for the resulting Excel file.
            Dim result As String = "Result-RetrieveAndExtractDataToNewExcelFile.xlsx"

            ' Saves the new workbook to a file with the specified file name in the Excel 2013 format.
            newBook.SaveToFile(result, ExcelVersion.Version2013)
            ' Release the resources used by the workbook
            workbook.Dispose()
            newBook.Dispose()

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
