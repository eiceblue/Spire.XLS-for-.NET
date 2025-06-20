Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace GetCellDisplayedText
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            'Access the first worksheet in the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            'Access the cell range B8 in the worksheet.
            Dim cell As CellRange = worksheet.Range("B8")
            'Set the numeric value of the cell to 0.012345.
            cell.NumberValue = 0.012345

            'Access the style of the cell.
            Dim style As CellStyle = cell.Style
            'Set the number format of the cell to "0.00".
            style.NumberFormat = "0.00"

            'Retrieve the value of the cell as a string.
            Dim cellValue As String = cell.Value

            'Retrieve the displayed text of the cell.
            Dim displayedText As String = cell.DisplayedText

            'Create a StringBuilder object to store the output content.
            Dim content As New StringBuilder()

            'Create the result string with the formatted values.
            Dim result As String = String.Format("B8 Value: " & cellValue & vbCrLf & "B8 displayed text: " & displayedText)

            'Append the result string to the StringBuilder.
            content.AppendLine(result)

            'Specify the output file name.
            Dim outputFile As String = "Output.txt"

            'Write the content of the StringBuilder to the specified text file.
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
