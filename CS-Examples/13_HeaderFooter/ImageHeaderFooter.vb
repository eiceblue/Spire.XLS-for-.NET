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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ImageHeaderFooter.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create an Image object from a specified image file
            Dim image As Image = image.FromFile("..\..\..\..\..\..\Data\Logo.png")

            ' Set the left header image and text for the page setup
            sheet.PageSetup.LeftHeaderImage = image
            sheet.PageSetup.LeftHeader = "&G"

            ' Set the center footer image and text for the page setup
            sheet.PageSetup.CenterFooterImage = image
            sheet.PageSetup.CenterFooter = "&G"

            ' Set the view mode of the sheet to Layout
            sheet.ViewMode = ViewMode.Layout

            ' Specify the output file name
            Dim result As String = "Output_ImageHeaderFooter.xlsx"

            ' Save the modified workbook to a specified path with Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
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
