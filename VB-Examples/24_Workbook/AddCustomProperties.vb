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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load a workbook from a specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AddCustomProperties.xlsx")

            ' Add a custom document property "_MarkAsFinal" with a value of True
            workbook.CustomDocumentProperties.Add("_MarkAsFinal", True)

            ' Add custom document properties with their respective names and values
            workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue")
            workbook.CustomDocumentProperties.Add("Phone number", 81705109)
            workbook.CustomDocumentProperties.Add("Revision number", 7.12)
            workbook.CustomDocumentProperties.Add("Revision date", Date.Now)

            ' Save the modified workbook to a new file with the specified file format (Version2013)
            workbook.SaveToFile("AddCustomProperties_result.xlsx", FileFormat.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
