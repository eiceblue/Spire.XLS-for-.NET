Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SetXlsSheetCenterOnPage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load an Excel file from a specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            ' Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the PageSetup object of the worksheet.
            Dim pageSetup As PageSetup = sheet.PageSetup

            ' Set the CenterHorizontally property to true, centering the sheet horizontally on the page.
            pageSetup.CenterHorizontally = True

            ' Set the CenterVertically property to true, centering the sheet vertically on the page.
            pageSetup.CenterVertically = True

            ' Specify the name for the result file.
            Dim result As String = "Result-SetXlsSheetCenterOnPage.xlsx"

            ' Save the modified workbook to a file with the specified result name and Excel version.
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
