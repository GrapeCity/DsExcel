Namespace Features.Worksheets
    Public Class UnprotectWorksheetWithPassword
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\\Medical office start-up expenses 1.xlsx")
            workbook.Open(fileStream)

            ' Use a password to protect all the worksheet. 
            ' If you forget the password, you cannot unprotect the worksheet.
            For Each worksheet In workbook.Worksheets
                worksheet.Protect("Y6dh!et5")
            Next

            ' Use the correct password to remove the above protection from the worksheet.
            For Each worksheet In workbook.Worksheets
                worksheet.Unprotect("Y6dh!et5")
            Next
        End Sub

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\\Medical office start-up expenses 1.xlsx"}
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
