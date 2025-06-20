Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace ManipulateTextBox
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook instance
            Dim workbook As New Workbook()

            ' Load the workbook from a specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ManipulateTextBoxControl.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Get the first TextBox control from the worksheet
            Dim tb As ITextBox = sheet.TextBoxes(0)

            ' Set the text of the TextBox
            tb.Text = "Spire.XLS for .NET"

            ' Set the horizontal alignment of the TextBox to center
            tb.HAlignment = CommentHAlignType.Center

            ' Set the vertical alignment of the TextBox to center
            tb.VAlignment = CommentVAlignType.Center

            ' Specify the output file name
            Dim output As String = "ManipulateTextBoxControl_out.xlsx"

            ' Save the modified workbook to the specified file path in Excel 2013 format
            workbook.SaveToFile(output, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
