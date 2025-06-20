Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace CopySheetToAnotherXlsFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a workbook.
            Dim workbook As New Workbook()

            'Get the first worksheet.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Put some data into header rows (A1:A4)
            For i As Integer = 1 To 5
                sheet.Range("A" & i).Text = String.Format("Header Row {0}", i)
                'sheet.Cells[i].Value = string.Format("Header Row {0}",i);
            Next i

            'Put some detail data (A5:A99)
            For i As Integer = 5 To 99
                sheet.Range("A" & i).Text = String.Format("Detail Row {0}", i)
                'sheet.Cells[i].Value = string.Format("Detail Row {0}",i);
            Next i

            'Define a pagesetup object based on the first worksheet.
            Dim pageSetup As PageSetup = sheet.PageSetup

            'The first five rows are repeated in each page. It can be seen in print preview.
            pageSetup.PrintTitleRows = "$1:$5"

            'Create another Workbook.
            Dim workbook1 As New Workbook()

            'Get the first worksheet in the book.
            Dim sheet1 As Worksheet = workbook1.Worksheets(0)

            'Copy worksheet to destination worsheet in another Excel file.
            sheet1.CopyFrom(sheet)

            Dim result As String = "Result-sourceFile.xlsx"
            Dim result1 As String = "Result-CopySheetToAnotherXlsFile.xlsx"

            'Save the source file we created.
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            'Save the destination file.
            workbook1.SaveToFile(result1, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ' Release the resources used by the workbook
            workbook1.Dispose()

            'Launch the MS Excel files.
            ExcelDocViewer(result)
			ExcelDocViewer(result1)
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
