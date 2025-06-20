Imports Spire.Xls
Imports System.IO
Imports System.Text

Namespace CheckAutoFitRowOrColumn
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Creates a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Creates a StringBuilder object to store the result.
            Dim result As New StringBuilder()

            'Loads an Excel file from the specified path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\CheckAutoFitRowsAndColumns.xlsx")

            'Checks if the second row in the worksheet has an auto fit row height.
            Dim isRowAutofit As Boolean = workbook.Worksheets(0).GetRowIsAutoFit(2)

            If isRowAutofit Then
                result.AppendLine("The second row is auto fit row height.")
            Else
                result.AppendLine("The second row is not auto fit row height.")
            End If

            'Checks if the second column in the worksheet has an auto fit column width.
            Dim isColAutofit As Boolean = workbook.Worksheets(0).GetColumnIsAutoFit(2)

            If isColAutofit Then
                result.AppendLine("The second column is auto fit column width.")
            Else
                result.AppendLine("The second column is not auto fit column width.")
            End If
            'Writes the result to a text file.
            File.WriteAllText("CheckAutoFitRowOrColumn_result.txt", result.ToString())
            ' Release the resources used by the workbook
            workbook.Dispose()

            FileViewer("CheckAutoFitRowOrColumn_result.txt")
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
