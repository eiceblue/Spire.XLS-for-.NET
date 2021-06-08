Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports Spire.Xls.Core.Spreadsheet.PivotTables
Imports System.Drawing.Imaging
Imports Spire.Xls.Core
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text

Namespace GetTextBoxByName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the default first  worksheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Insert a TextBox
			sheet.Range("A2").Text = "Name£º"
			Dim textBox As ITextBoxShape = sheet.TextBoxes.AddTextBox(2, 2, 18, 65)

			'Set the name 
			textBox.Name = "FirstTextBox"

			'Set string text for TextBox 
			textBox.Text = "Spire.XLS for .NET is a professional Excel .NET component that can be used to any type of .NET 2.0, 3.5, 4.0 or 4.5 framework application, both ASP.NET web sites and Windows Forms application."

			'Get the TextBox by the name
			Dim FindTextBox As ITextBoxShape = sheet.TextBoxes("FirstTextBox")

			'Get the TextBox text 
			Dim text As String = FindTextBox.Text

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set string format for displaying
			Dim result As String = String.Format("The text of """ & textBox.Name & """ is :" & text)

			'Add result string to StringBuilder
			content.AppendLine(result)

			'String for output file 
			Dim outputFile As String = "Output.txt"

			'Save them to a txt file
			File.WriteAllText(outputFile, content.ToString())

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
