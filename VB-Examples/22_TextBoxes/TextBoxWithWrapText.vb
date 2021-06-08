Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet.Shapes
Imports System.ComponentModel
Imports System.Text

Namespace TextBoxWithWrapText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk          
			workbook.LoadFromFile("..\..\..\..\..\..\Data\TextBoxSampleB.xlsx")

			Dim sheet As Worksheet = workbook.Worksheets(0)
			'Get the text box
			Dim shape As XlsTextBoxShape = TryCast(sheet.TextBoxes(0), XlsTextBoxShape)

			'Set wrap text
			shape.IsWrapText = True

			'Save the document
			Dim output As String = "TextBoxWithWrapText.xlsx"
			workbook.SaveToFile(output, ExcelVersion.Version2013)

			'View the document
			FileViewer(output)
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
