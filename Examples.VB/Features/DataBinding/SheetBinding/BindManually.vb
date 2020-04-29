Namespace Features.DataBinding.SheetBinding
    Public Class BindManually
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
#Region "Define custom classes"
            'Public Class SalesData
            '	Public Records As List(Of SalesRecord)
            'End Class

            'Public Class SalesRecord
            '	Public Area As String
            '	Public Salesman As String
            '	Public Product As String
            '	Public ProductType As String
            '	Public Sales As Integer
            'End Class
#End Region

#Region "Init data"
            Dim datasource = New SalesData With {.Records = New List(Of SalesRecord)()}

            Dim record1 = New SalesRecord With {
                .Area = "NorthChina",
                .Salesman = "Hellen",
                .Product = "Apple",
                .ProductType = "Fruit",
                .Sales = 120
            }
            datasource.Records.Add(record1)

            Dim record2 = New SalesRecord With {
                .Area = "NorthChina",
                .Salesman = "Hellen",
                .Product = "Banana",
                .ProductType = "Fruit",
                .Sales = 143
            }
            datasource.Records.Add(record2)

            Dim record3 = New SalesRecord With {
                .Area = "NorthChina",
                .Salesman = "Hellen",
                .Product = "Kiwi",
                .ProductType = "Fruit",
                .Sales = 322
            }
            datasource.Records.Add(record3)
#End Region

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Set AutoGenerateColumns as false
            worksheet.AutoGenerateColumns = False

            ' Bind columns manually.
            worksheet.Range("A:A").EntireColumn.BindingPath = "Area"
            worksheet.Range("B:B").EntireColumn.BindingPath = "Salesman"
            worksheet.Range("C:C").EntireColumn.BindingPath = "Product"
            worksheet.Range("D:D").EntireColumn.BindingPath = "ProductType"

            ' Set data source
            worksheet.DataSource = datasource.Records
        End Sub
    End Class
End Namespace
