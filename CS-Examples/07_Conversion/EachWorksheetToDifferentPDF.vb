Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace EachWorksheetToDifferentPDF
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\EachWorksheetToDifferentPDFSample.xlsx")

            ' Iterate through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets
                ' Set the file name for the PDF as the name of the current sheet
                Dim FileName As String = sheet.Name & ".pdf"

                ' Save the current sheet to PDF with the specified file name
                sheet.SaveToPdf(FileName)
            Next sheet
            ' Release the resources used by the workbook
            workbook.Dispose()

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
