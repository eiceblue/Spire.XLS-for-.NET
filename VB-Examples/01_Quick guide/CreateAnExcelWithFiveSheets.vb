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
			Dim start As Date = Date.Now
			'Create a workbook
			Dim workbook As New Workbook()
			workbook.CreateEmptySheets(5)
			For i As Integer = 0 To 4
				Dim sheet As Worksheet = workbook.Worksheets(i)
				sheet.Name = "Sheet" & i.ToString()
				For row As Integer = 1 To 150
					For col As Integer = 1 To 50
						sheet.Range(row, col).Text="row" & row.ToString() & " col" & col.ToString()
					Next col
				Next row
			Next i

			Dim result As String = "CreateAnExcelWithFiveSheets_result.xlsx"

			workbook.SaveToFile(result, ExcelVersion.Version2010)

			Dim [end] As Date = Date.Now
			Dim time As TimeSpan = [end].Subtract(start)
			MessageBox.Show("File has been created successfully! " & vbLf & "Time consumed (Seconds): " & time.TotalSeconds.ToString())

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
