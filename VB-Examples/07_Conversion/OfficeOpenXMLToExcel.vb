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
			Dim workbook As New Workbook()

			'Initialize worksheet
			Using fileStream As FileStream = File.OpenRead("..\..\..\..\..\..\Data\OfficeOpenXMLToExcel.Xml")
				workbook.LoadFromXml(fileStream)
			End Using
			workbook.SaveToFile("OfficeOpenXMLToExcel.xlsx", ExcelVersion.Version2010)

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
