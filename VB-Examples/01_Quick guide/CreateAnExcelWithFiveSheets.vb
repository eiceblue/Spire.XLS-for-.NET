Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CreateAnExcelWithFiveSheets
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Record the start time
            Dim start As Date = Date.Now

            ' Create a new Excel workbook object.
            Dim workbook As New Workbook()
            workbook.CreateEmptySheets(5)

            For i As Integer = 0 To 4
                ' Get the worksheet at the current index
                Dim sheet As Worksheet = workbook.Worksheets(i)
                ' Set the worksheet name 
                sheet.Name = "Sheet" & i.ToString()

                ' Fill in data by iterating through rows and columns
                For row As Integer = 1 To 150
                    For col As Integer = 1 To 50
                        sheet.Range(row, col).Text = "row" & row.ToString() & " col" & col.ToString() ' Set text content in the cell
                    Next col
                Next row
            Next i

            ' Specify the file name for the result
            Dim result As String = "CreateAnExcelWithFiveSheets_result.xlsx"

            ' Save the workbook to a file
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Record the end time
            Dim [end] As Date = Date.Now

            ' Calculate the elapsed time
            Dim time As TimeSpan = [end].Subtract(start)

            MessageBox.Show("File has been created successfully! " & vbLf & "Time consumed (Seconds): " & time.TotalSeconds.ToString())

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
