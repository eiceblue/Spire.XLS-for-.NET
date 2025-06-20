Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ReplaceTextInTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReplaceTextInTextBox.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Define the tag and replace strings
            Dim tag As String = "TAG_1$TAG_2"
            Dim replace As String = "Spire.XLS for .NET$Spire.XLS for JAVA"

            ' Iterate through each tag and replace its corresponding value in the worksheet
            For i As Integer = 0 To tag.Split("$"c).Length - 1
                ReplaceTextInTextBox(sheet, "<" & tag.Split("$"c)(i) & ">", replace.Split("$"c)(i))
            Next i

            ' Specify the output filename
            Dim output As String = "ReplaceTextInTextBox_out.xlsx"

            ' Save the modified workbook to a file in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
		End Sub
		Private Sub ReplaceTextInTextBox(ByVal sheet As Worksheet, ByVal sFind As String, ByVal sReplace As String)
			For i As Integer = 0 To sheet.TextBoxes.Count - 1
				Dim tb As ITextBox = sheet.TextBoxes(i)
				If Not String.IsNullOrEmpty(tb.Text) Then
					If tb.Text.Contains(sFind) Then
						tb.Text = tb.Text.Replace(sFind, sReplace)
					End If
				End If
			Next i
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
