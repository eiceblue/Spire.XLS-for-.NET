Imports System
Imports System.Windows.Forms
Imports Spire.Xls
Imports System.Drawing.Imaging

Namespace SpecificCellsToImage

	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			'Create a workbook
			Dim workbook As New Workbook()

			'Load the document from disk
			workbook.LoadFromFile("..\..\..\..\..\..\Data\ConversionSample1.xlsx")

			'Get the first worksheet in Excel file
			Dim sheet As Worksheet = workbook.Worksheets(0)

			'Specify Cell Ranges and Save to certain Image formats
			sheet.ToImage(1, 1, 7, 5).Save("image1.png", ImageFormat.Png)
			sheet.ToImage(8, 1, 15, 5).Save("image2.jpg", ImageFormat.Jpeg)
			sheet.ToImage(17, 1, 23, 5).Save("image3.bmp", ImageFormat.Bmp)

			'////////////////Use the following code for netstandard dlls/////////////////////////
'            
'            FileStream fileStream1 = new FileStream("SpecificCellsToImage1.png", FileMode.Create, FileAccess.Write);
'            sheet.ToImage(1, 1, 7, 5).CopyTo(fileStream1, 100);
'            FileStream fileStream2 = new FileStream("SpecificCellsToImage2.jpg", FileMode.Create, FileAccess.Write);
'            sheet.ToImage(8, 1, 15, 5).CopyTo(fileStream2, 100);
'            FileStream fileStream3 = new FileStream("SpecificCellsToImage3.bmp", FileMode.Create, FileAccess.Write);
'            sheet.ToImage(17, 1, 23, 5).CopyTo(fileStream3, 100);
'            fileStream1.Flush();
'            fileStream1.Close();
'            fileStream2.Flush();
'            fileStream2.Close();
'            fileStream3.Flush();
'            fileStream3.Close();
'			

			' Dispose of the workbook object to release resources
			workbook.Dispose()
		End Sub


		Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
			Close()
		End Sub

	End Class
End Namespace
