Namespace Features.PageSetup
    Public Class ConfigPrintTailRows
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("RepeatTails.xlsx")
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(1)

            'Repeat the row 67th at the bottom of each page when saving pdf.
            worksheet.PageSetup.PrintTailRows = "$50:$50"
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "RepeatTails.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\RepeatTails.xlsx"}
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
    End Class
End Namespace
