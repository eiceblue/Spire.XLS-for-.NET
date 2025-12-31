Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.IO
Imports Spire.Xls
Imports Spire.Xls.Charts

Namespace CopyPicture
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			' Create a new workbook
			Dim workbook As New Workbook()

			' Load an existing Excel document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

			' Get the first worksheet in the workbook
			Dim sheet1 As Worksheet = workbook.Worksheets(0)

			' Add a new worksheet as the destination sheet
			Dim destinationSheet As Worksheet = workbook.Worksheets.Add("DestSheet")

			' Get the first picture from the first worksheet
			Dim sourcePicture As ExcelPicture = sheet1.Pictures(0)

			' Add the image into the added worksheet at cell (2, 2)
			destinationSheet.Pictures.Add(2, 2, sourcePicture.Picture)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            //Get the image
'            SkiaSharp.SKBitmap image = sourcePicture.Picture;
'            //Add the image into the added worksheet 
'            destinationSheet.Pictures.Add(2, 2, image);
'            

			' Specify the output file name
			Dim outputFile As String = "Output.xlsx"

			' Save the modified workbook to the specified file using Excel 2013 format
			workbook.SaveToFile(outputFile, ExcelVersion.Version2013)

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			'Launching the output file.
			Viewer(outputFile)
		End Sub
		Private Sub Viewer(ByVal fileName As String)
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
