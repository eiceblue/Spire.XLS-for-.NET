Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing.Imaging
Imports Spire.Xls

Namespace ChartToImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "ChartToImage.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToImage.xlsx")

            ' Save the chart from the first worksheet of the workbook as an image and assign it to the 'image' variable
            Dim image As Image = workbook.SaveChartAsImage(workbook.Worksheets(0), 0)

            ' Save the image to a file named "Output.png" in PNG format
            image.Save("Output.png", ImageFormat.Png)
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            ExcelDocViewer("Output.png")
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
