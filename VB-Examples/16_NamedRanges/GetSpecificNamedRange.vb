Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core
Imports System.Text
Imports System.IO

Namespace GetSpecificNamedRange
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

            ' Get the name of the named range at index 1 in the workbook
            Dim name1 As String = workbook.NameRanges(1).Name

            ' Append information about accessing the specific named range by index to the StringBuilder
            sb.Append("Get the specific named range " & name1 & " by index" & vbCrLf)

            ' Get the name of the named range with the name "NameRange3" in the workbook
            Dim name2 As String = workbook.NameRanges("NameRange3").Name

            ' Append information about accessing the specific named range by name to the StringBuilder
            sb.Append("Get the specific named range " & name2 & " by name" & vbCrLf)

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
