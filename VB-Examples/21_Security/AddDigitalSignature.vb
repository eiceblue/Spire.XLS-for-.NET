Imports Spire.Xls
Imports Spire.Xls.Core.MergeSpreadsheet.Interfaces
Imports System.Security.Cryptography.X509Certificates

Namespace AddDigitalSignature
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()
		End Sub
		Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnRun.Click
            ' Create a new Workbook object
            Dim workbook As New Workbook()

            ' Load the Excel file from the specified path
            workbook.LoadFromFile("..\..\..\..\..\..\Data\DigitalSignature.xlsx")

            ' Specify the path of the input PFX file
            Dim inputFile_pfx As String = "..\..\..\..\..\..\Data\gary.pfx"

            ' Create a new X509Certificate2 object using the PFX file and password
            Dim cert As New X509Certificate2(inputFile_pfx, "e-iceblue")

            ' Specify the date and time for the digital signature
            Dim certtime As New Date(2020, 7, 1, 7, 10, 36)

            ' Add a digital signature to the Workbook using the specified certificate, reason, and time
            Dim dsc As IDigitalSignatures = workbook.AddDigitalSignature(cert, "e-iceblue", certtime)

            ' Specify the name of the resulting Excel file after adding the digital signature
            Dim result As String = "AddDigitalSignature.xlsx"

            ' Save the Workbook to the specified path in Excel 2010 format
            workbook.SaveToFile(result, ExcelVersion.Version2010)

            ' Release the resources used by the workbook
            workbook.Dispose()
            'View the document
            FileViewer(result)
		End Sub

		Private Sub FileViewer(ByVal fileName As String)
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
