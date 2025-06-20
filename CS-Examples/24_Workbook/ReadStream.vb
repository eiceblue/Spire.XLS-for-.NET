Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace ReadStream

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Open the file stream to read the Excel file
            Dim fileStream As FileStream = File.OpenRead("..\..\..\..\..\..\Data\ReadStream.xlsx")

            ' Set the file stream position to the beginning
            fileStream.Seek(0, SeekOrigin.Begin)

            ' Load the workbook from the file stream
            workbook.LoadFromStream(fileStream)

            ' Save the workbook to a new file named "ReadStream_result.xlsx" in Excel 2013 format
            workbook.SaveToFile("ReadStream_result.xlsx", ExcelVersion.Version2013)

            ' Release the resources used by the workbook
            workbook.Dispose()

            ' Release the resources used by  fileStream
            fileStream.Close()
            fileStream.Dispose()

            ExcelDocViewer("ReadStream_result.xlsx")
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
