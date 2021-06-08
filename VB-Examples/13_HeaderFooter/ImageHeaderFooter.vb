Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace ImageHeaderFooter
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Load a Workbook from disk
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ImageHeaderFooter.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Load an image from disk
			Dim image As Image = Image.FromFile("..\..\..\..\..\..\Data\Logo.png")

			'Set the image header
			sheet.PageSetup.LeftHeaderImage = image
			sheet.PageSetup.LeftHeader = "&G"

			'Set the image footer
			sheet.PageSetup.CenterFooterImage = image
			sheet.PageSetup.CenterFooter = "&G"

			'Set the view mode of the sheet
			sheet.ViewMode = ViewMode.Layout

			Dim result As String="Output_ImageHeaderFooter.xlsx"
			'Save and Launch
			workbook.SaveToFile(result, ExcelVersion.Version2010)
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
