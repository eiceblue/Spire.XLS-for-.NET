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
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load a workbook from a specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AccessDocumentProperties.xlsx")

            ' Create a StringBuilder object to store the extracted document properties
            Dim builder As New StringBuilder()

            ' Get the custom document properties of the workbook
            Dim properties As ICustomDocumentProperties = workbook.CustomDocumentProperties

            ' Retrieve and append the value of the "Editor" property to the StringBuilder
            Dim property1 As DocumentProperty = CType(properties("Editor"), DocumentProperty)
            builder.AppendLine(property1.Name.ToString() & " " & property1.Value.ToString())

            ' Retrieve and append the value of the first property in the collection to the StringBuilder
            Dim property2 As DocumentProperty = CType(properties(0), DocumentProperty)
            builder.AppendLine(property2.Name.ToString() & " " & property2.Value.ToString())

            ' Specify the output file name
            Dim result As String = "AccessDocumentProperties_out.txt"

            ' Write the contents of the StringBuilder to the output file
            File.WriteAllText(result, builder.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

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
