Imports Spire.Xls
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Namespace ChartToEMFImage
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file "ChartToEMFImage.xlsx" from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartToEMFImage.xlsx")

            ' Create a new MemoryStream object to store the EMF image data
            Using stream As New MemoryStream()
                ' Save the chart from the first worksheet of the workbook as an EMF image to the memory stream
                workbook.SaveChartAsEmfImage(workbook.Worksheets(0), 0, stream)

                ' Write the contents of the memory stream to a file named "EmfImage.emf"
                File.WriteAllBytes("EmfImage.emf", stream.ToArray())
            End Using
            ' Release the resources used by the workbook
            workbook.Dispose()

            'Launch the file
            ExcelDocViewer("EmfImage.emf")
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
