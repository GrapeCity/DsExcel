Namespace Features.Worksheets
    Public Class MoveWorksheet
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Load template file AgingReport.xlsx from resource
            Dim fileStream = GetResourceStream("AgingReport.xlsx")
            workbook.Open(fileStream)

            'Move the active sheet to the end of current workbook
            Dim move_worksheet = workbook.ActiveSheet.Move()
            move_worksheet.Name = "Move of " & workbook.ActiveSheet.Name
        End Sub

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "AgingReport.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\AgingReport.xlsx"}
            End Get
        End Property
    End Class
End Namespace
