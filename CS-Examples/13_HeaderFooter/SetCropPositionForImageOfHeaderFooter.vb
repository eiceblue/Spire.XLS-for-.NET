Imports Spire.Xls
Imports System.IO

Namespace SetCropPositionForImageOfHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			' Create a new instance of the Workbook class
			Dim workbook As New Workbook()

			' Load the workbook from the specified file path
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ImageInHeaderFooter.xlsx")

			' Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Set the cropping values for the left header picture
			sheet.PageSetup.LeftHeaderPictureCropTop = 0.2f
			sheet.PageSetup.LeftHeaderPictureCropBottom = 0.3f
			sheet.PageSetup.LeftHeaderPictureCropLeft = 0.3f
			sheet.PageSetup.LeftHeaderPictureCropRight = 0.2f

			' Set the cropping values for the left footer picture
			sheet.PageSetup.LeftFooterPictureCropTop = 0.2f
			sheet.PageSetup.LeftFooterPictureCropBottom = 0.3f
			sheet.PageSetup.LeftFooterPictureCropLeft = 0.3f
			sheet.PageSetup.LeftFooterPictureCropRight = 0.2f

			' Set the cropping values for the center header picture
			sheet.PageSetup.CenterHeaderPictureCropTop = 0.3f
			sheet.PageSetup.CenterHeaderPictureCropBottom = 0.4f
			sheet.PageSetup.CenterHeaderPictureCropLeft = 0.4f
			sheet.PageSetup.CenterHeaderPictureCropRight = 0.3f

			' Set the cropping values for the center footer picture
			sheet.PageSetup.CenterFooterPictureCropTop = 0.3f
			sheet.PageSetup.CenterFooterPictureCropBottom = 0.4f
			sheet.PageSetup.CenterFooterPictureCropLeft = 0.4f
			sheet.PageSetup.CenterFooterPictureCropRight = 0.3f

			' Set the cropping values for the right header picture
			sheet.PageSetup.RightHeaderPictureCropTop = 0.2f
			sheet.PageSetup.RightHeaderPictureCropBottom = 0.3f
			sheet.PageSetup.RightHeaderPictureCropLeft = 0.9f
			sheet.PageSetup.RightHeaderPictureCropRight = 0.4f

			' Set the cropping values for the right footer picture
			sheet.PageSetup.RightFooterPictureCropTop = 0.2f
			sheet.PageSetup.RightFooterPictureCropBottom = 0.3f
			sheet.PageSetup.RightFooterPictureCropLeft = 0.9f
			sheet.PageSetup.RightFooterPictureCropRight = 0.4f

			' Save the workbook to the specified file path with the specified file format
			Dim result As String = "result.xlsx"
			workbook.SaveToFile(result, FileFormat.Version2013)

			' Dispose workbook object
			workbook.Dispose()

			' Launch the file
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
