Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace OpenExistingFile
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz2.xlsx")

			'Add a new sheet, named MySheet
			Dim sheet As Worksheet = workbook.Worksheets.Add("MySheet")

			'Get the reference of "A1" cell from the cells collection of a worksheet
			sheet.Range("A1").Text = "Hello World"


			Dim result As String = "OpenExistingFile_result.xlsx"

			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			'View the document
			FileViewer(result)
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
