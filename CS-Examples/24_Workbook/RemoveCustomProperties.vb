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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\templateAz.xlsx")

            ' Get the custom document properties of the workbook
            Dim customDocumentProperties As ICustomDocumentProperties = workbook.CustomDocumentProperties

            ' Remove the custom document property named "Editor"
            customDocumentProperties.Remove("Editor")

            ' Specify the file name for the resulting workbook to be saved
            Dim result As String = "RemoveCustomProperties_result.xlsx"

            ' Save the modified workbook to a file with the specified name and Excel version (Version2010)
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
