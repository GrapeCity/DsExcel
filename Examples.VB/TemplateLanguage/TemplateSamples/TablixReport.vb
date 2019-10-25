Imports System.Data

Namespace Templates.TemplateSamples
    Public Class TablixReport
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_TablixReport.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim datasource As New DataTable
            With datasource.Columns
                .Add(New DataColumn("OrderID", GetType(Integer)))
                .Add(New DataColumn("Product", GetType(String)))
                .Add(New DataColumn("Sales", GetType(Double)))
                .Add(New DataColumn("ProductType", GetType(String)))
                .Add(New DataColumn("Year", GetType(String)))
                .Add(New DataColumn("Season", GetType(String)))
            End With

            With datasource.Rows
                .Add(1, "Röd Kaviar", 300, "Seafood", "2017", "Q3")
                .Add(2, "Spegesild", 144, "Seafood", "2017", "Q3")
                .Add(3, "Carnarvon Tigers", 600, "Seafood", "2017", "Q3")
                .Add(4, "Spegesild", 288, "Seafood", "2017", "Q4")
                .Add(5, "Carnarvon Tigers", 4250, "Seafood", "2017", "Q4")
                .Add(6, "Escargots de Bourgogne", 636, "Seafood", "2017", "Q4")
                .Add(7, "Röd Kaviar", 240, "Seafood", "2018", "Q1")
                .Add(8, "Carnarvon Tigers", 450, "Seafood", "2018", "Q1")
                .Add(9, "Röd Kaviar", 735, "Seafood", "2018", "Q2")
                .Add(10, "Røgede sild", 1377, "Seafood", "2018", "Q2")
                .Add(11, "Röd Kaviar", 1020, "Seafood", "2018", "Q3")
                .Add(12, "Røgede sild", 190, "Seafood", "2018", "Q3")
                .Add(13, "Röd Kaviar", 1725, "Seafood", "2018", "Q4")
                .Add(14, "Carnarvon Tigers", 3562, "Seafood", "2018", "Q4")
                .Add(15, "Sir Rodney's Marmalade", 4276, "Confections", "2017", "Q3")
                .Add(16, "Maxilaku", 880, "Confections", "2017", "Q3")
                .Add(17, "Maxilaku", 1040, "Confections", "2017", "Q4")
                .Add(18, "NuNuCa Nuß-Nougat-Creme", 716.8, "Confections", "2017", "Q4")
                .Add(19, "Sir Rodney's Marmalade", 2592, "Confections", "2018", "Q1")
                .Add(20, "Maxilaku", 1296, "Confections", "2018", "Q1")
                .Add(21, "Pavlova", 1473.4, "Confections", "2018", "Q1")
                .Add(22, "Sir Rodney's Marmalade", 4374, "Confections", "2018", "Q2")
                .Add(23, "Maxilaku", 1004, "Confections", "2018", "Q2")
                .Add(24, "Pavlova", 3075, "Confections", "2018", "Q2")
                .Add(25, "Sir Rodney's Marmalade", 1071, "Confections", "2018", "Q3")
                .Add(26, "Maxilaku", 860, "Confections", "2018", "Q3")
                .Add(27, "Pavlova", 732, "Confections", "2018", "Q3")
                .Add(28, "Sir Rodney's Marmalade", 1071, "Confections", "2018", "Q4")
                .Add(29, "Pavlova", 2634, "Confections", "2018", "Q4")
                .Add(30, "Sir Rodney's Scones", 1790, "Confections", "2018", "Q4")
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
                Return "Template_TablixReport.xlsx"
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
                Return New String() {"xlsx\Template_TablixReport.xlsx"}
            End Get
        End Property
    End Class


End Namespace
