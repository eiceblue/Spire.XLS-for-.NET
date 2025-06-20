Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls

Namespace SaveStream
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load an existing Excel file into the workbook
            workbook.LoadFromFile("..\..\..\..\..\..\Data\SaveStream.xls")

            ' Create a new FileStream object to write the workbook to a file
            Dim fileStream As New FileStream("SaveStream.xlsx", FileMode.Create)

            ' Save the workbook to the file stream in Excel 2010 format
            workbook.SaveToStream(fileStream, FileFormat.Version2010)

            ' Close the file stream
            fileStream.Close()

            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("SaveStream.xlsx")
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
