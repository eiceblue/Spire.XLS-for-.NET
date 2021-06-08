Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace AddCustomProperties
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook and load a file
			Dim workbook As New Workbook()
			workbook.LoadFromFile("..\..\..\..\..\..\Data\AddCustomProperties.xlsx")

			'Add a custom property to make the document as final
			workbook.CustomDocumentProperties.Add("_MarkAsFinal", True)

			'Add other custom properties to the workbook
			workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue")
			workbook.CustomDocumentProperties.Add("Phone number", 81705109)
			workbook.CustomDocumentProperties.Add("Revision number", 7.12)
			workbook.CustomDocumentProperties.Add("Revision date", Date.Now)

			'Save the document and launch it
			workbook.SaveToFile("AddCustomProperties_result.xlsx", FileFormat.Version2013)
			ExcelDocViewer("AddCustomProperties_result.xlsx")
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
