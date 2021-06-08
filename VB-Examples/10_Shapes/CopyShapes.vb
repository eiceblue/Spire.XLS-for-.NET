Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CopyShapes
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create line shape
			Dim line = sheet.TypedLines.AddLine()
			line.Top = 50
			line.Left = 30
			line.Width = 30
			line.Height = 50
			line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowDiamond
			line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

			Dim CopyShapes As Worksheet = workbook.Worksheets(1)
			'Copy the line into other sheet
			CopyShapes.TypedLines.AddCopy(line)

			'Create a button and then copy into other sheet
			Dim button = sheet.TypedRadioButtons.Add(5, 5, 20, 20)
			CopyShapes.TypedRadioButtons.AddCopy(button)

			'Create a textbox and then copy into other sheet
			Dim textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100)
			CopyShapes.TypedTextBoxes.AddCopy(textbox)

			'Create a checkbox and then copy into other sheet
			Dim checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20)
			CopyShapes.TypedCheckBoxes.AddCopy(checkbox)

			'Create a comboboxes and then copy into other sheet
			sheet.Range("A14").Value = "1"
			sheet.Range("A15").Value = "2"
			Dim ComboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30)
			ComboBoxes.ListFillRange = sheet.Range("A14:A15")
			CopyShapes.TypedComboBoxes.AddCopy(ComboBoxes)

			workbook.SaveToFile("CopyShapes.xlsx",ExcelVersion.Version2010)
			'View the document
			FileViewer("CopyShapes.xlsx")
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
