Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace CopyCellFormat
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel file from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Retrieve the first worksheet from the workbook..
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Get the number of rows in the worksheet.
            Dim count As Integer = sheet.Rows.Length
            'Iterate through each row.
            For i As Integer = 1 To count
                'Copy the cell style from column 2 (cell B) and apply it to the corresponding cell in column 5 (cell E).
                sheet.Range(String.Format("E{0}", i)).Style = sheet.Range(String.Format("B{0}", i)).Style
            Next i
            'Specify the file name for the output file.
            Dim result As String = "Result-CopyCellFormat.xlsx"

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
