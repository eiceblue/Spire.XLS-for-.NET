Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts
Imports System.Text

Namespace TraverseCellsValue
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CellValues.xlsx")

            'Retrieves the first worksheet from the workbook.
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            'Retrieves the collection of all cells in the worksheet.
            Dim cellRangeCollection() As CellRange = worksheet.Cells

            'Creates a new instance of the StringBuilder class for storing the output.
            Dim content As New StringBuilder()
            'Appends a line to the StringBuilder object.
            content.AppendLine("Values of the first sheet:")

            'Iterates through each cell range in the collection.
            For Each cellRange As CellRange In cellRangeCollection
                'Formats the cell address and value into a string.
                Dim result As String = String.Format("Cell: " & cellRange.RangeAddress & "   Value: " & cellRange.Value)

                'Appends the formatted string to the StringBuilder object.
                content.AppendLine(result)
            Next cellRange

            'Specifies the name of the output file.
            Dim outputFile As String = "Output.txt"

            'Writes the contents of the StringBuilder object to the specified file.
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
