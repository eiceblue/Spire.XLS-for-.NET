Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace AddListBoxControl
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

			'Set text for cells 
			sheet.Range("A7").Text = "Beijing"
			sheet.Range("A8").Text = "New York"
			sheet.Range("A9").Text = "ChengDu"
			sheet.Range("A10").Text = "Paris"
			sheet.Range("A11").Text = "Boston"
			sheet.Range("A12").Text = "London"

			sheet.Range("C13").Text = "City :"
			sheet.Range("C13").Style.Font.IsBold = True

			'Add listbox control
			Dim listBox As IListBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80)
			listBox.SelectionType = SelectionType.Single
			listBox.SelectedIndex = 2
			listBox.Display3DShading = True
			listBox.ListFillRange = sheet.Range("A7:A12")

			'Save the document
			Dim output As String = "InsertListBoxControl_out.xlsx"
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
