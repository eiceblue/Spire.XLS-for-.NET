Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Text
Imports System.IO

Namespace GetAllNamedRange
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new StringBuilder object
            Dim sb As New StringBuilder()

            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AllNamedRanges.xlsx")

            ' Get the collection of named ranges in the workbook
            Dim ranges As INameRanges = workbook.NameRanges

            ' Iterate through each named range and append its name to the StringBuilder
            For Each nameRange As INamedRange In ranges
                sb.Append(nameRange.Name & vbCrLf)
            Next nameRange

            ' Define the output file name as "result.txt"
            Dim result As String = "result.txt"

            ' Write the contents of the StringBuilder to a text file
            File.WriteAllText(result, sb.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

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
