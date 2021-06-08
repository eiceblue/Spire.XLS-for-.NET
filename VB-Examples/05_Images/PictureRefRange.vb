Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace PictureRefRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\PictureRefRange.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			sheet.Range("A1").Value = "Spire.XLS"
			sheet.Range("B3").Value = "E-iceblue"

			'Get the first picture in worksheet
			Dim picture As ExcelPicture = sheet.Pictures(0)

			'Set the reference range of the picture to A1:B3
			picture.RefRange = "A1:B3"

			'Save the Excel file
			Dim result As String = "PictureRefRange_out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the Excel file
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
