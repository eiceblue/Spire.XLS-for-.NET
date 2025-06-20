Imports Spire.Xls
Imports System.Text

Namespace MoveChartsheet
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the workbook from a file using a relative path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\MoveChartsheet.xlsx")

            ' Move the first chartsheet in the workbook to index 2 (moves it to the third position)
            workbook.Chartsheets(0).MoveSheet(2)

            ' Move the first chartsheet in the workbook to the beginning (index 0)
            workbook.Chartsheets(0).MoveChartsheet(0)

            ' Specify the output file name for saving the modified workbook
            Dim result As String = "MoveChartSheetResult.xlsx"

            ' Save the workbook to a file with the specified output name and Excel version
            workbook.SaveToFile(result, ExcelVersion.Version2013)

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
