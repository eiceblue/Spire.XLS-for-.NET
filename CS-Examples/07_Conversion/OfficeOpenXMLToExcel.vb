Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace OfficeOpenXMLToExcel
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Open the XML file as a FileStream for reading
            Using fileStream As FileStream = File.OpenRead("..\..\..\..\..\..\Data\OfficeOpenXMLToExcel.Xml")
                ' Load the XML data into the workbook
                workbook.LoadFromXml(fileStream)
            End Using

            ' Save the workbook to an Excel file in Excel 2010 format
            workbook.SaveToFile("OfficeOpenXMLToExcel.xlsx", ExcelVersion.Version2010)
            ' Release the resources used by the workbook
            workbook.Dispose()

            ExcelDocViewer("OfficeOpenXMLToExcel.xlsx")
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
