Imports Spire.Xls

Namespace CompressPictures
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\CompressPictures.xlsx")

			'Get the first worksheet
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			For Each sheet As Worksheet In workbook.Worksheets
				For Each picture As ExcelPicture In sheet.Pictures
					picture.Compress(50)
				Next picture
			Next sheet

			'Save to file
			Dim result As String = "CompressPictures_result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
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
