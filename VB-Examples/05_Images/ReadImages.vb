Imports System.Data.OleDb
Imports System.Collections
Imports System.ComponentModel

Imports Spire.Xls

Namespace ReadImages

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            'Create a new instance of the Workbook class.
            Dim workbook As New Workbook()

            'Load an existing Excel document from the specified file path.
            workbook.LoadFromFile("..\..\..\..\..\..\Data\ReadImages.xlsx")

            'Retrieve the first worksheet from the workbook.
            Dim sheet As Worksheet = workbook.Worksheets(0)

            'Access the first picture in the worksheet.
            Dim pic As ExcelPicture = sheet.Pictures(0)


			Using frm1 As New Form()
				Dim pic1 As New PictureBox()
				pic1.Image = pic.Picture
				frm1.Width = pic.Picture.Width
				frm1.Height = pic.Picture.Height
				frm1.StartPosition = FormStartPosition.CenterParent
				pic1.Dock = DockStyle.Fill
				frm1.Controls.Add(pic1)
				frm1.ShowDialog()
			End Using
            ' Release the resources used by the workbook
            workbook.Dispose()
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
