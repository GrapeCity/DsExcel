Namespace ExcelTemplates
    Public Class SimpleInvoice
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Simple invoice.xlsx from resource
            Dim fileStream = GetResourceStream("Simple invoice.xlsx")
            workbook.Open(fileStream)

            Dim worksheet = workbook.ActiveSheet

            ' fill some new items
            worksheet.Range("E09:H09").Value = New Object() {"DD1-001", "Item 3", 5.6, 12}
            worksheet.Range("E10:H10").Value = New Object() {"DD2-001", "Item 3", 8.5, 14}
            worksheet.Range("E11:H11").Value = New Object() {"DD3-001", "Item 3", 9.6, 16}
        End Sub
        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Simple invoice.xlsx"
            End Get
        End Property
        Public Overrides ReadOnly Property HasTemplate As Boolean
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
