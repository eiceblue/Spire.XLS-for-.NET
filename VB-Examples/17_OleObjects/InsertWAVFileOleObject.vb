Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO
Imports Spire.Xls.Core

Namespace InsertWavFileOLEObject
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Get the first sheet
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Add OLE object
			Dim oleObject As IOleObject = sheet.OleObjects.Add("..\..\..\..\..\..\Data\WAVFileSample.wav", Image.FromFile("..\..\..\..\..\..\Data\SpireXls.png"), OleLinkType.Embed)
			'Set the object location
			oleObject.Location = sheet.Range("B4")
			'Set the object type as package
			oleObject.ObjectType = OleObjectType.Package

			'Save and launch result file
			Dim result As String = "result.xlsx"
			workbook.SaveToFile(result, ExcelVersion.Version2010)
			ExcelDocViewer(result)

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
