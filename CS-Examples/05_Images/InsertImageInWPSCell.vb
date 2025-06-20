Imports Spire.Xls

Namespace InsertImageInWPSCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new workbook object
			Dim workbook As New Workbook()

			' Get the first worksheet
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			' Embed an image into the first cell
			  worksheet.Cells(0).InsertOrUpdateCellImage("..\..\..\..\..\..\Data\SpireXls.png", True)

			' Specify the filename for the resulting Excel file
			Dim result As String = "output.xlsx"

			' Save the workbook to the specified file in Excel 2010 format
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object
			workbook.Dispose()

			' View the document using a file viewer
			FileViewer(result)

			Me.Close()
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
