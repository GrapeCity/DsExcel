Namespace Features.PageSetup
    Public Class ConfigPrintTailColumns
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim fileStream = GetResourceStream("RepeatTails.xlsx")
            workbook.Open(fileStream)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            'Repeat the column "I" at the right of each page when saving pdf.
            worksheet.PageSetup.PrintTailColumns = "$I:$I"
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
