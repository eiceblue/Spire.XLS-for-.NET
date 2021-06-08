Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace SetTabColor
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook and load a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SetTabColor.xlsx")

			'Set the tab color of first sheet to be red 
			Dim worksheet As Worksheet = workbook.Worksheets(0)
			worksheet.TabColor = Color.Red

			'Set the tab color of first sheet to be green 
			worksheet = workbook.Worksheets(1)
			worksheet.TabColor = Color.Green

			'Set the tab color of first sheet to be blue 
			worksheet = workbook.Worksheets(2)
			worksheet.TabColor = Color.LightBlue

			'Save the document and launch it
			workbook.SaveToFile("SetTabColor_result.xlsx",ExcelVersion.Version2010)
			ExcelDocViewer("SetTabColor_result.xlsx")
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
