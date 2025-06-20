Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls
Imports System.IO

Namespace ToSVG
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object to represent an Excel workbook
            Dim workbook As New Workbook()

            ' Load an existing Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ToSVG.xlsx")

            ' Iterate through each worksheet in the workbook
            For i As Integer = 0 To workbook.Worksheets.Count - 1

                ' Create a FileStream object with the filename for the SVG file
                Dim fs As New FileStream(String.Format("sheet{0}.svg", i), FileMode.Create)

                ' Convert the current worksheet to SVG format and save it to the FileStream
                workbook.Worksheets(i).ToSVGStream(fs, 0, 0, 0, 0)

                ' Flush the FileStream to ensure all data is written
                fs.Flush()

                ' Close the FileStream to release resources
                fs.Close()
            Next i
            ' Release the resources used by the workbook
            workbook.Dispose()

            Process.Start("sheet0.svg")
		End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
			Close()
		End Sub
	End Class
End Namespace
