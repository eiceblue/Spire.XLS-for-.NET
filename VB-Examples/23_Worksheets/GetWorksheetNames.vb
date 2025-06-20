Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace GetWorksheetNames

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of the Workbook class
            Dim workbook As New Workbook()

            ' Load the workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\WorksheetSample3.xlsx")

            ' Create a StringBuilder object to store the result
            Dim stringBuilder As New StringBuilder()

            ' Iterate through each worksheet in the workbook
            For Each sheet As Worksheet In workbook.Worksheets

                ' Append the name of the current worksheet to the StringBuilder
                stringBuilder.AppendLine(sheet.Name)
            Next sheet

            ' Specify the output file name
            Dim output As String = "GetWorksheetNames.txt"

            ' Write the content of the StringBuilder to the output file
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
