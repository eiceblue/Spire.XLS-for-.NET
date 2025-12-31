Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Spire.Xls
Imports System.IO
Imports Spire.Xls.Core

Namespace InsertWavFileOLEObject
	Partial Public Class Form1
		Inherits Form

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			' Create a new workbook
			Dim workbook As New Workbook()

			' Get the first worksheet in the workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Add an OLE object
			Dim oleObject As IOleObject = sheet.OleObjects.Add("..\..\..\..\..\..\Data\WAVFileSample.wav", Image.FromFile("..\..\..\..\..\..\Data\SpireXls.png"), OleLinkType.Embed)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream fs = new FileStream(@"..\..\..\..\..\..\Data\SpireXls.png", FileMode.Open, FileAccess.Read, FileShare.Read);
'            byte[] bytes = new byte[fs.Length];
'            fs.Read(bytes, 0, bytes.Length);
'            fs.Close();
'            Stream ImgFile = new MemoryStream(bytes);
'            IOleObject oleObject = sheet.OleObjects.Add(@"..\..\..\..\..\..\Data\WAVFileSample.wav", ImgFile, OleLinkType.Embed);
'            

			' Set the location for the OLE object
			oleObject.Location = sheet.Range("B4")

			' Set the type of the OLE object as a package
			oleObject.ObjectType = OleObjectType.Package

			' Specify the output file name for the result
			Dim result As String = "result.xlsx"

			' Save the modified workbook to the specified file using Excel 2010 format
			workbook.SaveToFile(result, ExcelVersion.Version2010)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
			ExcelDocViewer(result)

		End Sub
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
