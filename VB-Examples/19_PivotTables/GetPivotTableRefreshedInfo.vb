Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text
Imports Spire.Xls.Core.Spreadsheet.PivotTables

Namespace GetPivotTableRefreshedInfo
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from a specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\PivotTable.xlsx")

            ' Get the first worksheet in the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Retrieve the PivotTable object from the worksheet and cast it as an XlsPivotTable
            Dim pivotTable As XlsPivotTable = TryCast(worksheet.PivotTables(0), XlsPivotTable)

            ' Get the refresh date and refreshed by information from the PivotTable's cache
            Dim dateTime As Date = pivotTable.Cache.RefreshDate
            Dim refreshedBy As String = pivotTable.Cache.RefreshedBy

            ' Create a StringBuilder object to store the content
            Dim content As New StringBuilder()

            ' Format the result with the refreshed by and refreshed date information
            Dim result As String = String.Format("Pivot table refreshed by: " & refreshedBy & vbCrLf & "Pivot table refreshed date: " & dateTime.ToString())

            ' Append the result to the content StringBuilder object
            content.AppendLine(result)

            ' Specify the output file name
            Dim outputFile As String = "Output.txt"

            ' Write the content of the StringBuilder object to the output file
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
