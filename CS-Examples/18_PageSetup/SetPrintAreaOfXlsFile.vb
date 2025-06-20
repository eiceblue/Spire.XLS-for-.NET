Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SetPrintAreaOfXlsFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the PageSetup object for the worksheet
            Dim pageSetup As PageSetup = sheet.PageSetup

            ' Set the print area for the worksheet to be printed
            pageSetup.PrintArea = "A1:E5"

            ' Specify the filename for the resulting Excel file
            Dim result As String = "Result-SetPrintAreaOfXlsFile.xlsx"

            ' Save the workbook to the specified file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
