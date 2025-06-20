Imports System.IO
Imports System.Text
Imports Spire.Xls

Namespace GetDefaultRowAndColumnCount

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Removes all existing worksheets from the workbook.
            workbook.Worksheets.Clear()

            'Creates a new empty worksheet and assigns it to the variable 'sheet'.
            Dim sheet As Worksheet = workbook.CreateEmptySheet()
            'Creates a new instance of the StringBuilder class.
            Dim sb As New StringBuilder()

            'Gets the number of rows in the worksheet.
            Dim rowCount As Integer = sheet.Rows.Length
            'Gets the number of columns in the worksheet.
            Dim columnCount As Integer = sheet.Columns.Length
            'Appends the row count information to the StringBuilder.
            sb.AppendLine("The default row count is: " & rowCount)
            'Appends the column count information to the StringBuilder.
            sb.AppendLine("The default column count is: " & columnCount)

            'Specifies the name of the output text file.
            Dim output As String = "GetDefaultRowAndColumnCount.txt"
            'Writes the content of the StringBuilder to the specified text file.
            File.WriteAllText(output, sb.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
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
