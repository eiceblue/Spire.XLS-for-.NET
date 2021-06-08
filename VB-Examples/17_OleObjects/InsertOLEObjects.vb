Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace InsertOLEObjects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'load Excel file
			Dim workbook As New Workbook()
			Dim ws As Worksheet = workbook.Worksheets(0)
			ws.Range("A1").Text = "Here is an OLE Object."
			'insert OLE object
			Dim xlsFile As String = "..\..\..\..\..\..\Data\InsertOLEObjects.xls"
			Dim image As Image = GenerateImage(xlsFile)
			Dim oleObject As IOleObject = ws.OleObjects.Add(xlsFile, image, OleLinkType.Embed)
			oleObject.Location = ws.Range("B4")
			oleObject.ObjectType = OleObjectType.ExcelWorksheet
			'save the file
			Dim result As String = "InsertOLEObjects_result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
			ExcelDocViewer(result)
		End Sub
		Private Function GenerateImage(ByVal fileName As String) As Image
			Dim book As New Workbook()
			book.LoadFromFile(fileName)
			book.Worksheets(0).PageSetup.LeftMargin = 0
			book.Worksheets(0).PageSetup.RightMargin = 0
			book.Worksheets(0).PageSetup.TopMargin = 0
			book.Worksheets(0).PageSetup.BottomMargin = 0
			Return book.Worksheets(0).ToImage(1, 1, 19, 5)
		End Function
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
