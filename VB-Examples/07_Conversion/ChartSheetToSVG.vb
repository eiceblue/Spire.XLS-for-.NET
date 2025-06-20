Imports Spire.Xls
Imports System.IO

Namespace ChartSheetToSVG

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new instance of Workbook
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ChartSheet.xlsx")

            ' Get the ChartSheet named "Chart1" from the workbook
            Dim cs As ChartSheet = workbook.GetChartSheetByName("Chart1")

            ' Specify the output file name for the SVG conversion
            Dim output As String = "ToSVG.svg"

            ' Create a new FileStream with the specified output file name
            Dim fs As New FileStream(String.Format(output), FileMode.Create)

            ' Convert the ChartSheet to SVG and write it to the FileStream
            cs.ToSVGStream(fs)

            ' Flush any buffered data in the FileStream
            fs.Flush()

            ' Close the FileStream
            fs.Close()
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
