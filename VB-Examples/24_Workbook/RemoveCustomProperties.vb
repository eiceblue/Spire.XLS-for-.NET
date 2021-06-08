Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text
Imports Spire.Xls.Core

Namespace RemoveCustomProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load a excel document
			workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

			'Retrieve a list of all custom document properties of the Excel file
			Dim customDocumentProperties As ICustomDocumentProperties = workbook.CustomDocumentProperties

			'Remove "Editor" custom document property
			customDocumentProperties.Remove("Editor")

			Dim result As String = "RemoveCustomProperties_result.xlsx"
			'Save to file
			workbook.SaveToFile(result, ExcelVersion.Version2010)
			'View the document
			ExcelDocViewer(result)
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
