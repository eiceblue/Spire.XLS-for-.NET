Imports Spire.Xls
Imports System.IO

Namespace GetFreezePaneRange
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook class
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\GetFreezePaneRange.xlsx")

            ' Get the first worksheet in the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Declare variables for storing row and column indices
            Dim rowIndex As Integer
            Dim colIndex As Integer

            ' Get the freeze pane indices of the worksheet
            sheet.GetFreezePanes(rowIndex, colIndex)

            ' Create a string representation of the freeze pane range
            Dim range As String = "Row index: " & rowIndex & ", column index: " & colIndex

            ' Define the output file name
            Dim result As String = "GetFreezePaneCellRange_result.txt"

            ' Write the freeze pane range to the output file
            File.WriteAllText(result, range)

            ' Release the resources used by the workbook
            workbook.Dispose()

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

		Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

		End Sub
	End Class
End Namespace
