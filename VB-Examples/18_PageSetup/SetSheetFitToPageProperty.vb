Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet

Namespace SetSheetFitToPageProperty
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of Workbook
            Dim workbook As New Workbook()

            'Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_4.xlsx")

            'Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Set the FitToPagesTall property of the PageSetup object to 1, which means fit to 1 page tall
            sheet.PageSetup.FitToPagesTall = 1

            'Set the FitToPagesWide property of the PageSetup object to 1, which means fit to 1 page wide
            sheet.PageSetup.FitToPagesWide = 1

            'Specify the filename for the resulting workbook that will be saved
            Dim result As String = "Result-SetSheetFitToPageProperty.xlsx"

            'Save the workbook to the specified file path in Excel 2013 format
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
