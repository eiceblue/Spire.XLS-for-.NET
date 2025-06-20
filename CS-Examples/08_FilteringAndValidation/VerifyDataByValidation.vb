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
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file named "Sample.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Sample.xlsx")

            ' Get the first worksheet (index 0) from the workbook
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Specify the range of cell B4 and assign it to the variable "cell"
            Dim cell As CellRange = worksheet.Range("B4")

            ' Get the data validation settings for the cell
            Dim validation As Validation = cell.DataValidation

            ' Parse the minimum and maximum values from the data validation formulas
            Dim minimum As Double = Double.Parse(validation.Formula1)
            Dim maximum As Double = Double.Parse(validation.Formula2)

            ' Create a StringBuilder object to store the validation results
            Dim content As New StringBuilder()

            ' Iterate from 5 to 99 with a step of 40
            For i As Integer = 5 To 99 Step 40
                ' Set the cell value to the current iteration value
                cell.NumberValue = i

                ' Declare a string variable to store the result message
                Dim result As String = Nothing

                ' Check if the cell value is outside the valid range
                If cell.NumberValue < minimum OrElse cell.NumberValue > maximum Then
                    result = String.Format("Is input " & i & " a valid value for this Cell: false")
                Else
                    result = String.Format("Is input " & i & " a valid value for this Cell: true")
                End If

                ' Append the result message to the StringBuilder object
                content.AppendLine(result)
            Next i

            ' Specify the output filename for the validation results
            Dim outputFile As String = "Output.txt"

            ' Write the content of the StringBuilder object to a text file
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
