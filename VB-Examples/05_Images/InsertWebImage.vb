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
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim URL As String = "http://www.e-iceblue.com/downloads/demo/Logo.png"

			'Instantiate the web client object
			Dim webClient As New WebClient()

			'Extract image data into memory stream
			Dim objImage As MemoryStream = New System.IO.MemoryStream(webClient.DownloadData(URL))

			Dim image As Image = Image.FromStream(objImage)

			'Add the image in the sheet
			sheet.Pictures.Add(3, 2, image)

			'Save and launch result file
			Dim result As String = "result.xlsx"
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
