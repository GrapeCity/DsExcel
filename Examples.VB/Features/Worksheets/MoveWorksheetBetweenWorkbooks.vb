Namespace Features.Worksheets
    Public Class MoveWorksheetBetweenWorkbooks
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Load template file Home inventory.xlsx from resource to the source workbook
            Dim source_workbook As New Excel.Workbook
            Dim source_fileStream = GetResourceStream("Home inventory.xlsx")
            source_workbook.Open(source_fileStream)

            'Move content of active sheet from source workbook to the current workbook before the first sheet
            Dim move_worksheet = source_workbook.ActiveSheet.MoveBefore(workbook.Worksheets(0))
            move_worksheet.Name = "Move of Home inventory"
            move_worksheet.Activate()

            workbook.Theme = source_workbook.Theme
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Home inventory.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Home inventory.xlsx"}
            End Get
        End Property
    End Class
End Namespace
