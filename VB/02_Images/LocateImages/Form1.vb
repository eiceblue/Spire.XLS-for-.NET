Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace LocateImages
	Partial Public Class Form1
		Inherits Form
		Public Sub New()

			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
				Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\LocateImages.xlsx")
			'Get the first sheet
				  Dim sheet As Worksheet = workbook.Worksheets(0)

			Dim pic As ExcelPicture = sheet.Pictures(0)
			pic.LeftColumnOffset = 300
			pic.TopRowOffset = 300

			'Save and Launch
				  workbook.SaveToFile("Output.xlsx",ExcelVersion.Version2010)
				  ExcelDocViewer(workbook.FileName)
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
