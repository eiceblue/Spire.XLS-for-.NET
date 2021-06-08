Imports Spire.Xls
Imports System.Drawing.Imaging
Imports System.IO

Namespace SheetToEMF

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

			'Get the first worksheet in excel workbook
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Create a memory stream
			Dim stream As New MemoryStream()

			'Save excel worksheet into EMF stream
			sheet.ToEMFStream(stream, 1, 1, 28, 8, EmfType.EmfPlusDual)

			'Save to file
			Dim img As Image = Image.FromStream(stream)
			Dim output As String = "ToEMF.emf"
			img.Save(output)

			'Launch the Excel file
			ExcelDocViewer(output)
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
