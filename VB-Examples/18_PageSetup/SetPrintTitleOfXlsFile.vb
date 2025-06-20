Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SetPrintTitleOfXlsFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the page setup settings for the worksheet
            Dim pageSetup As PageSetup = sheet.PageSetup

            ' Set the columns to repeat on each printed page
            pageSetup.PrintTitleColumns = "$A:$B"

            ' Set the rows to repeat on each printed page
            pageSetup.PrintTitleRows = "$1:$2"

            ' Specify the name of the resulting Excel file
            Dim result As String = "Result-SetPrintTitleOfXlsFile.xlsx"

            ' Save the modified workbook to a file in Excel 2013 format
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
