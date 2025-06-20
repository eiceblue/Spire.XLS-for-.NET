Imports Spire.Xls
Imports System.IO

Namespace GetImageLink
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\hyperlink.xlsx")

			'Get the first picture of the first worksheet
			Dim picture As ExcelPicture = workbook.Worksheets(0).Pictures(0)

			'Get the address
			Dim address As String = picture.GetHyperLink().Address

			' Write the address to the txt file
			Dim file As String = "address.txt"
			System.IO.File.WriteAllText(file, address)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			OutputViewer(file)

		End Sub
		Private Sub OutputViewer(ByVal fileName As String)
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
