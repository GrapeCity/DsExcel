Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class SaveMultipleWorkbooksToPDF
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("Any year calendar1.xlsx")
            workbook.Open(fileStream)

            Dim workbook2 As New Excel.Workbook
            Dim fileStream2 As Stream = GetResourceStream("Any year calendar (Ion theme)1.xlsx")
            workbook2.Open(fileStream2)

            'Create a printmanager.
            Dim printManager As New Excel.PrintManager

            'Save the workbook1 and workbook2 into one pdf file.
            printManager.SavePDF(outputStream, workbook, workbook2)
        End Sub

        Public Overrides ReadOnly Property SavePageInfos As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Any year calendar1.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Any year calendar1.xlsx", "xlsx\Any year calendar (Ion theme)1.xlsx"}
            End Get
        End Property
    End Class
End Namespace
