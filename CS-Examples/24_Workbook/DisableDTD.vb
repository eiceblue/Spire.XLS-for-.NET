Imports Spire.Xls
Imports System.IO

Namespace DisableDTD
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim outputFile_E As String = "Ex.txt"
			Try
				Dim outputFile As String = "DisableDTD.xlsx"
			' Create a new workbook object
			Dim workbook As New Workbook()

			' Disable DTD
			workbook.ProhibitDtd = True

			'Load the file from disk.
			 workbook.LoadFromFile("..\..\..\..\..\..\Data\haveDtd.xlsx")

			'Save
			 workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			' Dispose of the workbook object
			workbook.Dispose()

			 FileViewer(outputFile)
			Catch ex As Exception
				Dim stream As New FileStream(outputFile_E, FileMode.Append)
				Dim sw As New StreamWriter(stream)
				sw.WriteLine(ex.ToString() & "Disable DTD processing：" & ex.ToString())
				sw.Flush()
				sw.Close()
				FileViewer(outputFile_E)
			End Try

			Me.Close()

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
