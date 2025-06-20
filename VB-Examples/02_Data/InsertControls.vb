Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace InsertControls
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim wb As New Workbook()

            ' Load an existing workbook from a file named "InsertControls.xlsx"
            wb.LoadFromFile("..\..\..\..\..\..\Data\InsertControls.xlsx")

            ' Get the first worksheet from the workbook
            Dim ws As Worksheet = wb.Worksheets(0)

            ' Add a textbox to the worksheet at position (9, 2) with width 25 and height 100
            Dim textbox As ITextBoxShape = ws.TextBoxes.AddTextBox(9, 2, 25, 100)
            textbox.Text = "Hello World"

            ' Add a checkbox to the worksheet at position (11, 2) with width 15 and height 100
            Dim cb As ICheckBox = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100)
            cb.CheckState = Spire.Xls.CheckState.Checked
            cb.Text = "Check Box 1"

            ' Add a RadioButton to the worksheet at position (13, 2) with width 15 and height 100
            Dim rb As IRadioButton = ws.RadioButtons.Add(13, 2, 15, 100)
            rb.Text = "Option 1"

            ' Add a combobox to the worksheet at position (15, 2) with width 15 and height 100
            Dim cbx As IComboBoxShape = TryCast(ws.ComboBoxes.AddComboBox(15, 2, 15, 100), IComboBoxShape)
            cbx.ListFillRange = ws.Range("A41:A47")

            ' Save the workbook to a file named "Result.xlsx" in Excel 2010 format
            wb.SaveToFile("Result.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            wb.Dispose()

            ExcelDocViewer("Result.xlsx")
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
