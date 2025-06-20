Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace AccessCell
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\AccessCell.xlsx")
            'Creates a new StringBuilder object to store text.
            Dim builder As New StringBuilder()

            'Gets the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Accesses the cell with the name "A1".
            Dim range1 As CellRange = sheet.Range("A1")
            'Appends the value of range1 to the StringBuilder.
            builder.AppendLine("Value of range1: " & range1.Text)

            'Accesses the cell at row 2, column 1 (B1 in Excel).
            Dim range2 As CellRange = sheet.Range(2, 1)
            'Appends the value of range2 to the StringBuilder.
            builder.AppendLine("Value of range2: " & range2.Text)

            'Accesses the cell at row 2 in the default column (A2 in Excel).
            Dim range3 As CellRange = sheet.Cells(2)
            'Appends the value of range3 to the StringBuilder.
            builder.AppendLine("Value of range3: " & range3.Text)

            'Specifies the filename for the resulting text file.
            Dim result As String = "AccessCell_out.txt"
            'Writes the content of the StringBuilder to the text file.
            File.WriteAllText(result, builder.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the txt file
            OutputViewer(result)
		End Sub
		Private Sub OutputViewer(ByVal fileName As String)
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
