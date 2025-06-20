Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System.Text
Imports System.IO

Namespace RetrieveExternalFileHyperlinks
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new workbook object
            Dim workbook As New Workbook()

            ' Load an existing workbook from the specified file path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\RetrieveExternalFileHyperlinks.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a StringBuilder object to store the content
            Dim content As New StringBuilder()

            ' Iterate through each hyperlink in the worksheet
            For Each item As HyperLink In sheet.HyperLinks
                ' Retrieve the address of the hyperlink
                Dim address As String = item.Address

                ' Retrieve the name of the worksheet containing the hyperlink
                Dim sheetName As String = item.Range.WorksheetName

                ' Retrieve the range of cells associated with the hyperlink
                Dim range As CellRange = item.Range

                ' Add a formatted line to the content StringBuilder
                content.AppendLine(String.Format("Cell[{0},{1}] in sheet """ & sheetName & """ contains File URL: {2}", range.Row, range.Column, address))
            Next item

            ' Define the output file name as "Result-RetrieveExternalFileHyperlinks.txt"
            Dim result As String = "Result-RetrieveExternalFileHyperlinks.txt"

            ' Write the content of the StringBuilder to a text file
            File.WriteAllText(result, content.ToString())

            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the MS Excel file.
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
