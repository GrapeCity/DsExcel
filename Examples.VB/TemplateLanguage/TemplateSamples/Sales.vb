Imports System.Data

Namespace Templates.TemplateSamples
    Public Class Sales
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_Sales.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim datasource = New DataTable

            With datasource.Columns
                .Add(New DataColumn("Area", GetType(String)))
                .Add(New DataColumn("Salesman", GetType(String)))
                .Add(New DataColumn("Product", GetType(String)))
                .Add(New DataColumn("ProductType", GetType(String)))
                .Add(New DataColumn("Sales", GetType(Integer)))
            End With

            With datasource.Rows
                .Add("NorthChina", "Hellen", "Apple", "Fruit", 120)
                .Add("NorthChina", "Hellen", "Banana", "Fruit", 143)
                .Add("NorthChina", "Hellen", "Kiwi", "Fruit", 322)
                .Add("NorthChina", "Hellen", "Carrots", "Vegetable", 154)
                .Add("NorthChina", "Fancy", "Carrots", "Vegetable", 131)
                .Add("NorthChina", "Fancy", "Cabbage", "Vegetable", 98)
                .Add("NorthChina", "Fancy", "Potato", "Vegetable", 212)
                .Add("NorthChina", "Fancy", "Apple", "Fruit", 102)
                .Add("NorthChina", "Ivan", "Apple", "Fruit", 164)
                .Add("NorthChina", "Ivan", "Kiwi", "Fruit", 213)
                .Add("NorthChina", "Ivan", "Potato", "Vegetable", 56)
                .Add("NorthChina", "Ivan", "Cabbage", "Vegetable", 265)
                .Add("SouthChina", "Adam", "Cabbage", "Vegetable", 112)
                .Add("SouthChina", "Adam", "Carrots", "Vegetable", 354)
                .Add("SouthChina", "Adam", "Banana", "Fruit", 277)
                .Add("SouthChina", "Adam", "Apple", "Fruit", 105)
                .Add("SouthChina", "Bob", "Kiwi", "Fruit", 402)
                .Add("SouthChina", "Bob", "Banana", "Fruit", 133)
                .Add("SouthChina", "Bob", "Cabbage", "Vegetable", 252)
                .Add("SouthChina", "Bob", "Potato", "Vegetable", 265)
            End With
#End Region

            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property IsNew As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property TemplateName As String
            Get
                Return "Template_Sales.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property HasTemplate As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownloadZip As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources As String()
            Get
                Return New String() {"xlsx\Template_Sales.xlsx"}
            End Get
        End Property
    End Class
End Namespace
