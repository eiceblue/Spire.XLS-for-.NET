Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace GetWorksheetOfChart
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object.
            Dim workbook As New Workbook()

            ' Load the Excel file into the workbook.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

            ' Get the first worksheet in the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Get the first chart in the worksheet.
            Dim chart As Chart = worksheet.Charts(0)

            ' Cast the chart's worksheet as a Worksheet object.
            Dim wSheet As Worksheet = TryCast(chart.Worksheet, Worksheet)

            ' Create a StringBuilder object to store the content.
            Dim content As New StringBuilder()

            ' Create a string with the sheet name and the chart's sheet name.
            Dim result As String = String.Format("Sheet Name: " & worksheet.Name & vbCrLf & "Charts' sheet Name: " & wSheet.Name)

            ' Append the result string to the content StringBuilder.
            content.AppendLine(result)

            ' Specify the output file name.
            Dim outputFile As String = "Output.txt"

            ' Write the contents of the content StringBuilder to a text file.
            File.WriteAllText(outputFile, content.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launching the output file.
            Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
