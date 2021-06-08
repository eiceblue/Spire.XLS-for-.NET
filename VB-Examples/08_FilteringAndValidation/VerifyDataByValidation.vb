Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace VerifyDataByValidation
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

			'Get the specified data range
			Dim minimum As Double = Double.Parse(validation.Formula1)
			Dim maximum As Double = Double.Parse(validation.Formula2)

			'Create StringBuilder to save 
			Dim content As New StringBuilder()

			'Set different numbers for the cell
			For i As Integer = 5 To 99 Step 40
				cell.NumberValue = i
				Dim result As String=Nothing
				'Verify 
				If cell.NumberValue < minimum OrElse cell.NumberValue > maximum Then
					'Set string format for displaying
					result = String.Format("Is input " & i & " a valid value for this Cell: false")
				Else
					'Set string format for displaying
					result = String.Format("Is input " & i & " a valid value for this Cell: true")
				End If
				'Add result string to StringBuilder
				content.AppendLine(result)
			Next i
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
