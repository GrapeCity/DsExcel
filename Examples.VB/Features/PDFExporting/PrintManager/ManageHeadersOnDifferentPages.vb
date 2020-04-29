Imports System.IO

Namespace Features.PDFExporting.PrintManager
    Public Class ManageHeadersOnDifferentPages
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook, outputStream As MemoryStream)
            Dim fileStream As Stream = GetResourceStream("MultipleHeaders.xlsx")
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            Dim repeatSettings As IList(Of RepeatSetting) = New List(Of RepeatSetting)

            'The title rows of the "B2:F87" is "$2:$2"
            Dim repeatSetting As New RepeatSetting With {
                .TitleRowStart = 1,
                .TitleRowEnd = 1,
                .Range = worksheet.Range("B2:F87")
            }
            repeatSettings.Add(repeatSetting)

            'The title rows of the "B91:F189" is "$91:$91"
            Dim repeatSetting2 As New RepeatSetting With {
                .TitleRowStart = 88,
                .TitleRowEnd = 88,
                .Range = worksheet.Range("B89:F149")
            }
            repeatSettings.Add(repeatSetting2)

            'Create a PrintManager.
            Dim printManager As New Excel.PrintManager

            'Get the pagination information of the worksheet.
            Dim pages As IList(Of PageInfo) = printManager.Paginate(worksheet, Nothing, repeatSettings)

            'Save the pages into pdf file.
            printManager.SavePDF(outputStream, pages)
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

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "MultipleHeaders.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\MultipleHeaders.xlsx"}
            End Get
        End Property
    End Class
End Namespace
