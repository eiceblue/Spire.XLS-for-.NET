Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.Text
Imports System.IO

Namespace GetCellAddress
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object.
            Dim workbook As New Workbook()

            'Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ExcelSample_N1.xlsx")

            'Access the first worksheet in the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Create a StringBuilder object to store the output text.
            Dim builder As New StringBuilder()

            'Define a CellRange object that represents the range of cells A1 to B5 in the worksheet.
            Dim range As CellRange = sheet.Range("A1:B5")

            'Retrieve the local address of the range.
            Dim address As String = range.RangeAddressLocal
            builder.AppendLine("Address of range: " & address)

            'Get the total number of cells within the range.
            Dim count As Integer = range.CellsCount
            builder.AppendLine("Cell count of range: " & count.ToString())

            'Retrieve the local address of the entire column of the range.
            Dim entireColAddress As String = range.EntireColumn.RangeAddressLocal
            builder.AppendLine("Address of entire column of the range: " & entireColAddress)

            'Retrieve the local address of the entire row of the range.
            Dim entireRowAddress As String = range.EntireRow.RangeAddressLocal
            builder.AppendLine("Address of entire row of the range " & entireRowAddress)

            'Specify the output file name.
            Dim output As String = "GetCellAddress_out.txt"
            'Write the content of the StringBuilder to the specified text file.
            File.WriteAllText(output, builder.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the txt file
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
