Namespace Features.DataBinding
    Public Class CellBinding
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
#Region "Define custom classes"
            'Public Class SalesRecord
            '	Public Area As String
            '	Public Salesman As String
            '	Public Product As String
            '	Public ProductType As String
            '	Public Sales As Integer
            'End Class
#End Region

#Region "Init data"
            Dim record As New SalesRecord With {
                .Area = "NorthChina",
                .Salesman = "Hellen",
                .Product = "Apple",
                .ProductType = "Fruit",
                .Sales = 120
            }
#End Region

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Set binding path for cell.
            worksheet.Range("A1").BindingPath = "Area"
            worksheet.Range("B2").BindingPath = "Salesman"
            worksheet.Range("C2").BindingPath = "Product"
            worksheet.Range("D3").BindingPath = "ProductType"

            ' Set data source.
            worksheet.DataSource = record
        End Sub
    End Class

    Public Class SalesRecord
        Public Area As String
        Public Salesman As String
        Public Product As String
        Public ProductType As String
        Public Sales As Integer
    End Class
End Namespace
