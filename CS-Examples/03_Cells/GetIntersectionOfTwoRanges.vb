Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace GetIntersectionOfTwoRanges
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\Template_Xls_1.xlsx")

            'Retrieves the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Finds the intersection of two ranges: "A2:D7" and "B2:E8".
            Dim range As CellRange = sheet.Range("A2:D7").Intersect(sheet.Range("B2:E8"))

            'Appends a line to the StringBuilder object, specifying the intersection of the two ranges.
            Dim content As New StringBuilder()
            content.AppendLine("The intersection of the two ranges ""A2:D7"" and ""B2:E8"" is:")

            'Iterates through each cell range in the intersection.
            For Each r As CellRange In range
                'Appends the string representation of the cell value to the StringBuilder object.
                content.AppendLine(r.Value.ToString())
            Next r
            'Specifies the name of the output file.
            Dim result As String = "Result-GetTheIntersectionOfTwoRanges.txt"

            'Writes the contents of the StringBuilder object to the specified file.
            File.WriteAllText(result, content.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file.
            ExcelDocViewer(result)
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
