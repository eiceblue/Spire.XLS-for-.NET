Imports System.IO
Imports System.Text
Imports Spire.Xls
Imports Spire.Xls.Collections
Imports Spire.Xls.Core

Namespace GetProperties

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample1.xlsx")

            ' Get the built-in document properties of the workbook
            Dim properties1 As BuiltInDocumentProperties = workbook.DocumentProperties

            ' Create a StringBuilder object to store the extracted properties
            Dim stringBuilder As New StringBuilder()
            stringBuilder.AppendLine("Excel Properties:")

            ' Iterate over the built-in document properties and append their names and values to the StringBuilder
            For i As Integer = 0 To properties1.Count - 1
                Dim name As String = properties1(i).Name
                Dim value As String = properties1(i).Value.ToString()
                stringBuilder.AppendLine(name & ": " & value)
            Next i
            stringBuilder.AppendLine()

            ' Get the custom document properties of the workbook
            Dim properties2 As ICustomDocumentProperties = workbook.CustomDocumentProperties
            stringBuilder.AppendLine("Custom Properties:")

            ' Iterate over the custom document properties and append their names and values to the StringBuilder
            For i As Integer = 0 To properties2.Count - 1
                Dim name As String = properties2(i).Name
                Dim value As String = properties2(i).Value.ToString()
                stringBuilder.AppendLine(name & ": " & value)
            Next i

            ' Specify the output file name
            Dim output As String = "GetProperties.txt"

            ' Write the contents of the StringBuilder to the output file
            File.WriteAllText(output, stringBuilder.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the Excel file
            ExcelDocViewer(output)
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
