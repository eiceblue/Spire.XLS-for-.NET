Imports Spire.Xls
Imports System.ComponentModel
Imports System.Text

Namespace CreateFiftyExcelFiles
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim start As Date = Date.Now
			For n As Integer = 0 To 49
				Dim workbook As New Workbook()
				workbook.CreateEmptySheets(5)
				For i As Integer = 0 To 4
					Dim sheet As Worksheet = workbook.Worksheets(i)
					sheet.Name = "Sheet" & i.ToString()
					For row As Integer = 1 To 150
						For col As Integer = 1 To 50
							sheet.Range(row, col).Text = "row" & row.ToString() & " col" & col.ToString()
						Next col
					Next row
				Next i

				workbook.SaveToFile("Workbook" & n & ".xlsx", ExcelVersion.Version2010)
			Next n
			Dim [end] As Date = Date.Now
			Dim time As TimeSpan = [end].Subtract(start)
			MessageBox.Show("50 File(s) have been created successfully! " & vbLf & "Time consumed (Seconds): " & time.TotalSeconds.ToString())
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
