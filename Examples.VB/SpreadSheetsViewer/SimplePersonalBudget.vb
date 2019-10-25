Namespace SpreadSheetsViewer
    Public Class SimplePersonalBudget
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file
            Dim fileStream = GetResourceStream("Simple personal budget.xlsx")
            workbook.Open(fileStream)
        End Sub

        Public Overrides ReadOnly Property HasTemplate As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property IsViewReadOnly As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property ShowCode As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Simple personal budget.xlsx"
            End Get
        End Property


        Public Overrides ReadOnly Property CanDownloadZip As Boolean
            Get
                Return False
            End Get
        End Property
        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Simple personal budget.xlsx"}
            End Get
        End Property
    End Class
End Namespace
