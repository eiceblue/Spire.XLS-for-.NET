Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace DuplicateCellRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object..
            Dim workbook As New Workbook()

            'Load an existing Excel file from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Copy the data and formatting from the source range (A6:F6) to the destination range (A16:F16) in the same worksheet, while maintaining the format.
            sheet.Copy(sheet.Range("A6:F6"), sheet.Range("A16:F16"), True)

            'Specify the file name for the output file.
            Dim result As String = "Result-DuplicateCellRange.xlsx"

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
