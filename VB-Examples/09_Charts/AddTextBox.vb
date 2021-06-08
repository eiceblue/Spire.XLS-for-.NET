Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports Spire.Xls.Core

Namespace AddTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a Workbook
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AddTextBox.xlsx")

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Get the first chart
			Dim chart As Chart = sheet.Charts(0)

			'Add a Textbox
			Dim textbox As ITextBoxLinkShape = chart.Shapes.AddTextBox()
			textbox.Width = 1200
			textbox.Height = 320
			textbox.Left = 1000
			textbox.Top = 480
			textbox.Text = "This is a textbox"

			'Save and Launch
			workbook.SaveToFile("Output.xlsx", ExcelVersion.Version2010)
			ExcelDocViewer("Output.xlsx")
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
