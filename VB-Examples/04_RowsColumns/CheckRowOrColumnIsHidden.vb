Imports Spire.Xls
Imports System.IO
Imports System.Text

Namespace CheckRowOrColumnIsHidden
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Instantiate a new Workbook object to create a workbook.
            Dim workbook As New Workbook()

            'Load the Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CheckRowOrColumnIsHidden.xlsx")

            'Get the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Instantiate a new StringBuilder object to store the result.
            Dim result As New StringBuilder()
            'Specify the row index.
            Dim rowIndex As Integer = 2
            'Specify the column index.
            Dim columnIndex As Integer = 2
            'Check if the specified row is hidden using the GetRowIsHide method.
            Dim rowIsHide As Boolean = sheet.GetRowIsHide(rowIndex)

            If rowIsHide Then
                'Add a message to the result indicating that the second row is hidden.
                result.AppendLine("The second row is hidden.")
            Else
                'Add a message to the result indicating that the second row is not hidden.
                result.AppendLine("The second row is not hidden.")
            End If
            'Check if the specified column is hidden using the GetColumnIsHide method.
            Dim columnIsHide As Boolean = sheet.GetColumnIsHide(columnIndex)

            If columnIsHide Then
                'Add a message to the result indicating that the second column is hidden.
                result.AppendLine("The second column is hidden.")
            Else
                'Add a message to the result indicating that the second column is not hidden.
                result.AppendLine("The second column is not hidden.")
            End If
            'Write the result to a text file named "CheckRowOrColumnIsHidden_result.txt".
            File.WriteAllText("CheckRowOrColumnIsHidden_result.txt", result.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

        End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
			Try
				Process.Start(fileName)
			Catch
			End Try
		End Sub
	End Class
End Namespace
