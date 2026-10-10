Imports Spire.Xls
Imports System
Imports System.Windows.Forms

Namespace SkipHeaderLengthCheck
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Initialize a new Workbook instance
            Dim workbook As Workbook = New Workbook()

            ' Get the first worksheet in the workbook
            Dim worksheet As Worksheet = workbook.Worksheets[0]

            ' Disable the validation for header/footer length to bypass limit checks
            workbook.ConverterSetting.ValidateHeaderFooterLength = False

            ' Define a very long string to intentionally exceed the default header length limit
            Dim fullLine As String = "══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════"

            ' Set the long string as the left header content
            worksheet.PageSetup.LeftHeader = $"{fullLine}"

            ' Specify the output file name for the PDF
            Dim result As String = "SkipHeaderLengthCheck.pdf"

            ' Export the worksheet to a PDF file
            worksheet.SaveToPdf(result)

            ' Dispose of the workbook to release resources
            workbook.Dispose()

            ' View the document using a file viewer
            FileViewer(result)
        End Sub

        Private Sub FileViewer(ByVal fileName As String)
            Try
                System.Diagnostics.Process.Start(fileName)
        Catch
        End Try
        End Sub

        Private Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs)
            Close()
        End Sub
    End Class
End Namespace
