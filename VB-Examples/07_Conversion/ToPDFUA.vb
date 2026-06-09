Imports Spire.Xls
Imports System
Imports System.IO
Imports System.Windows.Forms

Namespace ToPDFUA
    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            'Create a workbook
            Dim workbook As Workbook = New Workbook()

            'Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\SampleB_2.xlsx")

            ' Convert excel to PDF/UA
            workbook.ConverterSetting.PdfConformanceLevel = Spire.Xls.Pdf.PdfConformanceLevel.Pdf_UA1

            ' Specify the output file name for the EXCEL result
            Dim result As String = "ToPDFUA_out.pdf"

            ' Save the workbook as a PDF file with the name "sample.pdf"
            workbook.SaveToFile(result, FileFormat.PDF)

            ' Dispose of the workbook object to release resources
            workbook.Dispose()

            'View the document
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

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)

        End Sub
    End Class
End Namespace
