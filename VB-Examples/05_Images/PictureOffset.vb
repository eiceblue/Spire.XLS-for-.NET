Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace PictureOffset
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

			'Insert a picture
			Dim pic As ExcelPicture = sheet.Pictures.Add(2, 2,"..\..\..\..\..\..\Data\logo.png")

			'Set left offset and top offset from the current range
			pic.LeftColumnOffset = 200
			pic.TopRowOffset = 100

			'Save the Excel file
			Dim result As String = "PictureOffset_out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the file
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
