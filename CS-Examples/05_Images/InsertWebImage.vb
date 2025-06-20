Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO
Imports System.Net

Namespace InsertWebImage
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new workbook object.
            Dim workbook As New Workbook()

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Specify the URL of the image to be downloaded.
            Dim URL As String = "http://www.e-iceblue.com/downloads/demo/Logo.png"

            'Create a WebClient object.
            Dim webClient As New WebClient()

            'Download the image from the specified URL and store it in a MemoryStream.
            Dim objImage As MemoryStream = New System.IO.MemoryStream(webClient.DownloadData(URL))

            'Create an Image object from the downloaded image data.
            Dim image As Image = image.FromStream(objImage)

            'Insert the image at cell coordinates (3, 2) in the worksheet.
            sheet.Pictures.Add(3, 2, image)

            'Specify the name of the resulting file after adding the image.
            Dim result As String = "result.xlsx"
            'Save the modified workbook to the specified output file using Excel 2010 format.
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
