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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AccessDocumentProperties.xlsx")

            ' Add a custom document property named "Test" with the value "MyNamedRange"
            workbook.CustomDocumentProperties.Add("Test", "MyNamedRange")

            ' Get the collection of custom document properties
            Dim properties As ICustomDocumentProperties = workbook.CustomDocumentProperties

            ' Retrieve the specific custom document property named "Test"
            Dim [property] As DocumentProperty = CType(properties("Test"), DocumentProperty)

            ' Set the LinkToContent property of the custom document property to True
            [property].LinkToContent = True

            ' Specify the output file name as "LinkToContentProperty_out.xlsx"
            Dim result As String = "LinkToContentProperty_out.xlsx"

            ' Save the modified workbook to a file in Excel 2013 format
            workbook.SaveToFile(result, ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

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
