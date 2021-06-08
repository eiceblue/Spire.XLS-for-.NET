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
			Dim wb As New Workbook()
			wb.LoadFromFile("..\..\..\..\..\..\Data\InsertControls.xlsx")
			Dim ws As Worksheet = wb.Worksheets(0)

			'Add a textbox 
			Dim textbox As ITextBoxShape = ws.TextBoxes.AddTextBox(9, 2, 25, 100)
			textbox.Text = "Hello World"
			'Add a checkbox 
			Dim cb As ICheckBox = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100)
			cb.CheckState = Spire.Xls.CheckState.Checked
			cb.Text = "Check Box 1"
			'Add a RadioButton 
			Dim rb As IRadioButton = ws.RadioButtons.Add(13, 2, 15, 100)
			rb.Text = "Option 1"

			'Add a combox
			Dim cbx As IComboBoxShape = TryCast(ws.ComboBoxes.AddComboBox(15, 2, 15, 100), IComboBoxShape)
			cbx.ListFillRange = ws.Range("A41:A47")

			wb.SaveToFile("Result.xlsx", ExcelVersion.Version2010)

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
