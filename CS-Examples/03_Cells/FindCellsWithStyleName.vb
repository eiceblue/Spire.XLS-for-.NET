Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace FindCellsWithStyleName
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new workbook object.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SampleB_2.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Retrieve the style name of the cell A1.
            Dim styleName As String = sheet.Range("A1").CellStyleName

            'Retrieve the range of cells that contain data in the worksheet.
            Dim ranges As CellRange = sheet.AllocatedRange
            'Iterate through each cell range in the allocated range.
            For Each cc As CellRange In ranges
                'Check if the current cell range has the same style name as the one retrieved from cell A1.
                If cc.CellStyleName = styleName Then
                    'Set the value of the current cell range to "Same style".
                    cc.Value = "Same style"
                End If
            Next cc
            'Specify the file name for the output file.
            Dim result As String = "FindCellsWithStyleName_result.xlsx"

            'Save the modified workbook to the specified output file using Excel 2010 version.
            workbook.SaveToFile(result, ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
