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
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

            ' Get the first worksheet from the workbook
            Dim sheet As Worksheet = workbook.Worksheets(0)

            ' Create a new MemoryStream to store the EMF stream
            Dim stream As New MemoryStream()

            ' Convert the worksheet to EMF format and save it to the MemoryStream
            sheet.ToEMFStream(stream, 1, 1, 28, 8, EmfType.EmfPlusDual)

            ' Create an Image object from the MemoryStream
            Dim img As Image = Image.FromStream(stream)

            ' Specify the output file name for saving the EMF image
            Dim output As String = "ToEMF.emf"

            ' Save the EMF image to a file
            img.Save(output)
            ' Release the resources used by the workbook
            workbook.Dispose()

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
