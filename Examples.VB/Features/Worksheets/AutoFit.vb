Namespace Features.Worksheets
    Public Class AutoFit
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Auto fit column width of range 'A1'
            With worksheet.Range!A1
                .Value = "Grapecity Documents for Excel"
                .Columns.AutoFit()
            End With

            'Auto fit row height of range 'B2'
            With worksheet.Range!B2
                .Value = "Grapecity"
                .Font.Size = 20
                .Rows.AutoFit()
            End With

            'Auto fit column width and row height of range 'C3'
            With worksheet.Range!C3
                .Value = "Grapecity Documents for Excel"
                .Font.Size = 32
                .AutoFit()
            End With
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property
    End Class
End Namespace
