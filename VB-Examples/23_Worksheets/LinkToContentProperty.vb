Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace LinkToContentProperty
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AccessDocumentProperties.xlsx")

			'Add a custom document property
			workbook.CustomDocumentProperties.Add("Test", "MyNamedRange")

			'Get the added document property
			Dim properties As ICustomDocumentProperties = workbook.CustomDocumentProperties
			Dim [property] As DocumentProperty = CType(properties("Test"), DocumentProperty)

			'Link to content 
			[property].LinkToContent = True

			'Save the document
			Dim result As String = "LinkToContentProperty_out.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2013)

			'Launch the Excel file
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
