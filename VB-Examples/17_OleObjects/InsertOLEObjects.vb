Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Spire.Xls
Imports Spire.Xls.Core

Namespace InsertOLEObjects
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			' Create a new workbook
			Dim workbook As New Workbook()

			' Get the first worksheet in the workbook
			Dim ws As Worksheet = workbook.Worksheets(0)

			' Set the text in cell A1
			ws.Range("A1").Text = "Here is an OLE Object."

			' Insert an OLE object
			Dim xlsFile As String = "..\..\..\..\..\..\Data\InsertOLEObjects.xls"
			Dim image As Image = GenerateImage(xlsFile)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'              
'            Stream image = GenerateImage(xlsFile);
'            

			Dim oleObject As IOleObject = ws.OleObjects.Add(xlsFile, image, OleLinkType.Embed)
			oleObject.Location = ws.Range("B4")
			oleObject.ObjectType = OleObjectType.ExcelWorksheet

			' Specify the output file name for the result
			Dim result As String = "InsertOLEObjects_result.xlsx"

			' Save the modified workbook to the specified file using Excel 2010 format
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
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
				System.Diagnostics.Process.Start(fileName)
			Catch
			End Try
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub


	End Class
End Namespace
