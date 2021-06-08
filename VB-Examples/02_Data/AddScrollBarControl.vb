Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddScrollBarControl
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

			'Get the first worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Set a value for range B10
			sheet.Range("B10").Value2 = 1
			sheet.Range("B10").Style.Font.IsBold = True

			'Add scroll bar control
			Dim scrollBar As IScrollBarShape = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20)
			scrollBar.LinkedCell = sheet.Range("B10")
			scrollBar.Min = 1
			scrollBar.Max = 150
			scrollBar.IncrementalChange = 1
			scrollBar.Display3DShading = True

			'Save the document
			Dim output As String = "AddScrollBarControl_out.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'Launch the Excel file
			ExcelDocViewer(output)
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
