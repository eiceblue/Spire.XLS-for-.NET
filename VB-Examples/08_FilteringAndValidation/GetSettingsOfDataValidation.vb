Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace GetSettingsOfDataValidation
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "Sample.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Specify the cell B4 for which we want to retrieve the validation settings
            Dim cell As CellRange = worksheet.Range("B4")

            ' Get the data validation settings for the specified cell
            Dim validation As Validation = cell.DataValidation

            ' Retrieve various settings from the validation object
            Dim allowType As String = validation.AllowType.ToString()
            Dim compareOperator As String = validation.CompareOperator.ToString()
            Dim formula1 As String = validation.Formula1.ToString()
            Dim formula2 As String = validation.Formula2.ToString()
            Dim ignoreBlank As String = validation.IgnoreBlank.ToString()

            ' Create a StringBuilder object to save the output content
            Dim content As New StringBuilder()

            ' Set the string format for displaying the validation settings
            Dim result As String = String.Format("Settings of Validation: " & vbCrLf & "Allow Type: " & allowType & vbCrLf & "Compare Operator: " & compareOperator & vbCrLf & "Formula 1: " & formula1 & vbCrLf & "Formula 2: " & formula2 & vbCrLf & "Ignore Blank: " & ignoreBlank)

            ' Add the result string to the StringBuilder
            content.AppendLine(result)

            ' Specify the output file name as "Output.txt"
            Dim outputFile As String = "Output.txt"

            ' Save the contents of the StringBuilder to a text file
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
