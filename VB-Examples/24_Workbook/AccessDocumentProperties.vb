Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Text
Imports System.IO

Namespace AccessDocumentProperties
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

			'Create string builder
			Dim builder As New StringBuilder()

			'Get all document properties
			Dim properties As ICustomDocumentProperties = workbook.CustomDocumentProperties

			'Access document property by property name
			Dim property1 As DocumentProperty = CType(properties("Editor"), DocumentProperty)
			builder.AppendLine(property1.Name & " " & property1.Value)

			'Access document property by property index
			Dim property2 As DocumentProperty = CType(properties(0), DocumentProperty)
			builder.AppendLine(property2.Name & " " & property2.Value)

			'Save to txt file
			Dim result As String = "AccessDocumentProperties_out.txt"
			File.WriteAllText(result, builder.ToString())

			'Launch the file 
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
