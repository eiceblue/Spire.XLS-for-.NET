Imports System
Imports System.Windows.Forms
Imports Spire.Xls
Imports System.Text
Imports System.IO
Imports System.Drawing

Namespace GetEmbeddedImages

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

	   Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
			' Create a new Workbook instance
			Dim wb As New Workbook()

			' Load the Excel document from a specific file path
			wb.LoadFromFile("..\..\..\..\..\..\Data\EmbedImageViaWps.xlsx")

			' Access the first worksheet in the workbook
			Dim sheet As Worksheet = wb.Worksheets(0)

			' Retrieve an array of Excel pictures from the worksheet
			Dim pc() As ExcelPicture = sheet.CellImages

			' Iterate through each Excel picture in the array
			For i As Integer = 0 To pc.Length - 1
				Dim ep As ExcelPicture = pc(i)
				Dim image As Image = ep.Picture

				' Save the image as a PNG file with a unique name based on the index
				image.Save("result-" & i & ".png", System.Drawing.Imaging.ImageFormat.Png)

				'////////////////Use the following code for netstandard dlls///////////////////////// 
'				               
'                Stream img = sheet.ToImage(0,0,0,0);
'                FileStream fileStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write);
'                img.CopyTo(fileStream, 100);
'                fileStream.Flush();
'                fileStream.Close();
'                img.Close();
'                
			Next i
	   End Sub

		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub

	End Class
End Namespace
