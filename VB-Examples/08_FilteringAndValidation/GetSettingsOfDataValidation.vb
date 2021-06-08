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
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample.xlsx")

			'Get first worksheet of the workbook
			Dim worksheet As Worksheet = workbook.Worksheets(0)

			'Cell B4 has the Decimal Validation
			Dim cell As CellRange = worksheet.Range("B4")

			'Get the valditation of this cell
			Dim validation As Validation = cell.DataValidation

			'Get the settings
			Dim allowType As String = validation.AllowType.ToString()
			Dim data As String = validation.CompareOperator.ToString()
			Dim minimum As String = validation.Formula1.ToString()
			Dim maximum As String = validation.Formula2.ToString()
			Dim ignoreBlank As String = validation.IgnoreBlank.ToString()

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set string format for displaying
			Dim result As String = String.Format("Settings of Validation: " & vbCrLf & "Allow Type: " & allowType & vbCrLf & "Data: " & data & vbCrLf & "Minimum: " & minimum &vbCrLf & "Maximum: " & maximum & vbCrLf & "IgnoreBlank: " & ignoreBlank)

			'Add result string to StringBuilder
			content.AppendLine(result)

			'String for output file 
			Dim outputFile As String = "Output.txt"

			'Save them to a txt file
			File.WriteAllText(outputFile, content.ToString())

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
