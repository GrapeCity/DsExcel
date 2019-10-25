Namespace Features.Worksheets
    Public Class Tag
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Add tag for worksheet
            worksheet.Tag = "This is a Tag for sheet."

            ' Add tag for range A1:B2
            worksheet.Range("A1:B2").Tag = "This is a Tag for A1:B2"

            ' Add tag for row 4
            worksheet.Range("A4").EntireRow.Tag = "This is a Tag for Row 4"

            ' Add tag for column F
            worksheet.Range("F5").EntireColumn.Tag = "This is a Tag for Column F"

            ' Note:
            ' If you are using Visual Basic, when setting a boxed value type to the Tag property,
            ' a copy of the value will be assigned.
        End Sub

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

        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
