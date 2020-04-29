Namespace Features.Workbook
    Public Class UnprotectWorkbook
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            'Open an excel file
            Dim fileStream = GetResourceStream("xlsx\\Medical office start-up expenses 1.xlsx")
            workbook.Open(fileStream)

            ' Protects the workbook with a password so that other users cannot view hidden worksheets, add, move, delete, hide, or rename worksheets.
            ' The protection only happens when you open it with an Excel application.
            workbook.Protect("Y6dh!et5")

            ' Use the correct password to remove the above protection from the workbook.
            workbook.Unprotect("Y6dh!et5")
        End Sub

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\\Medical office start-up expenses 1.xlsx"}
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownload As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
