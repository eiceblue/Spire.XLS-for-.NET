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
            ' Create a new workbook
            Dim workbook As New Workbook()

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Set the text in cell A2 to "Name"
            sheet.Range("A2").Text = "Name£º"

            ' Add a text box to the worksheet at position (2, 2) with width 18 and height 65
            Dim textBox As ITextBoxShape = sheet.TextBoxes.AddTextBox(2, 2, 18, 65)

            ' Set the name of the text box to "FirstTextBox"
            textBox.Name = "FirstTextBox"

            ' Set the text inside the text box
            textBox.Text = "Spire.XLS for .NET is a professional Excel .NET component that can be used in any type of .NET 2.0, 3.5, 4.0, or 4.5 framework application, both ASP.NET web sites and Windows Forms applications."

            ' Retrieve the text box with the name "FirstTextBox" from the worksheet
            Dim FindTextBox As ITextBoxShape = sheet.TextBoxes("FirstTextBox")

            ' Get the text from the retrieved text box
            Dim text As String = FindTextBox.Text

            ' Create a StringBuilder to store the content
            Dim content As New StringBuilder()

            ' Format a string with the text box name and its text
            Dim result As String = String.Format("The text of """ & textBox.Name & """ is: " & text)

            ' Append the formatted result to the content StringBuilder
            content.AppendLine(result)

            ' Specify the output file path
            Dim outputFile As String = "Output.txt"

            ' Write the content to the output file
            File.WriteAllText(outputFile, content.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

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
