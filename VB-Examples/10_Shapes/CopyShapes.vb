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
            ' Create a new workbook and initialize the worksheet
            Dim workbook As New Workbook()
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Add a typed line shape to the worksheet
            Dim line = sheet.TypedLines.AddLine()
            line.Top = 50
            line.Left = 30
            line.Width = 30
            line.Height = 50
            line.BeginArrowHeadStyle = ShapeArrowStyleType.LineArrowDiamond
            line.EndArrowHeadStyle = ShapeArrowStyleType.LineArrow

            ' Copy the line shape to another worksheet
            Dim CopyShapes As Worksheet = workbook.Worksheets(1)
            CopyShapes.TypedLines.AddCopy(line)

            ' Add a typed radio button shape to the worksheet
            Dim button = sheet.TypedRadioButtons.Add(5, 5, 20, 20)
            CopyShapes.TypedRadioButtons.AddCopy(button)

            ' Add a typed text box shape to the worksheet
            Dim textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100)
            CopyShapes.TypedTextBoxes.AddCopy(textbox)

            ' Add a typed check box shape to the worksheet
            Dim checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20)
            CopyShapes.TypedCheckBoxes.AddCopy(checkbox)

            ' Set values for specific cells in the worksheet
            sheet.Range("A14").Value = "1"
            sheet.Range("A15").Value = "2"

            ' Add a typed combo box shape to the worksheet and set its list fill range
            Dim ComboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30)
            ComboBoxes.ListFillRange = sheet.Range("A14:A15")
            CopyShapes.TypedComboBoxes.AddCopy(ComboBoxes)

            ' Save the workbook to a file named "CopyShapes.xlsx" using Excel 2010 format
            workbook.SaveToFile("CopyShapes.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()
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
