Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.IO

Imports Spire.Xls

Namespace ToOfficeOpenXML
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			Dim workbook As New Workbook()
			Dim sheet As Worksheet = workbook.Worksheets(0)
			sheet.Range("A1").Text = "Hello World"
			sheet.Range("B1").Style.KnownColor = ExcelColors.Gray25Percent
			sheet.Range("C1").Style.KnownColor= ExcelColors.Gold
			workbook.SaveAsXml("sample.xml")

			Process.Start(Path.Combine(Application.StartupPath,"Sample.xml"))
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
