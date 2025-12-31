Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Spire.Xls
Imports System.IO
Imports System.Net

Namespace InsertWebImage
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			' Create a new workbook.
			Dim workbook As New Workbook()

			' Get the first sheet from the workbook.
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Specify the URL of the image to be downloaded.
			Dim URL As String = "http://www.e-iceblue.com/downloads/demo/Logo.png"

			' Instantiate a web client object.
			Dim webClient As New WebClient()

			' Extract the image data into a memory stream.
			Dim objImage As MemoryStream = New System.IO.MemoryStream(webClient.DownloadData(URL))

			' Add the image to the worksheet at a specific location (row 3, column 2).
			sheet.Pictures.Add(3, 2, objImage)

			' Specify the resulting file name.
			Dim result As String = "result.xlsx"

			' Save the modified workbook to a file using Excel 2010 format.
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources.
			workbook.Dispose()

			' Launch the file
			ExcelDocViewer(result)
		End Sub
		Private Sub ExcelDocViewer(ByVal fileName As String)
			Try
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub
	End Class


End Namespace
