Imports Spire.Xls
Imports Spire.Xls.Core.Spreadsheet
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Namespace ToMarkdownExportOptions

    Public Partial Class Form1 : Inherits Form
        Public Sub New()
            InitializeComponent()
        End Sub

        Private Sub btnRun_Click(ByVal sender As Object, ByVal e As EventArgs)
            ' Create a new workbook
            Dim workbook As Workbook = New Workbook()

            'Load the document from disk
            workbook.LoadFromFile(@"..\..\..\..\..\..\Data\ToMarkdownExportOptions.xlsx")

            ' Create export options for Markdown format
            Dim options As MarkdownOptions = New MarkdownOptions()

            ' Set whether to save images with relative paths
            options.SavePicInRelativePath = True

            ' Set whether to save hyperlinks as Markdown reference format
            options.SaveHyperlinkAsRef = True

            ' Save the workbook as Markdown with the specified options
            Dim result As String = "ToMarkdownExportOptions_out.md"
            workbook.SaveToMarkdown(result, options)
            
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
