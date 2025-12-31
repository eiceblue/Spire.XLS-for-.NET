Imports System
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Spire.Xls

Namespace SheetToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\SheetToImage.xlsx")

			'Get the first worksheet in excel workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			' Save to image
			sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn).Save("SheetToImage.png")

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            Stream image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);
'            string filename = String.Format("SheetToImage.png");
'            FileStream fileStream = new FileStream(filename, FileMode.Create, FileAccess.Write);
'            image.CopyTo(fileStream, 100);
'            fileStream.Flush();
'            fileStream.Close();
'            image.Close();
'            

			' Dispose of the workbook object to release resources
			workbook.Dispose()

			' Launch the file
			ExcelDocViewer("SheetToImage.png")
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
