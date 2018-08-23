Imports System.IO
Namespace Features.Workbook
    Public Class ImportCsvFileToWorkbook
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim stream As Stream = GetTemplateStream()
            'Open csv file stream.
            workbook.Open(stream, OpenFileFormat.Csv)
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Information.csv"
            End Get
        End Property
        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return True
            End Get
        End Property
        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
