Imports System.Data

Namespace Templates.DataSource
    Public Class DataSet
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_Scores.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim team1 As New Data.DataTable
            With team1.Columns
                .Add(New DataColumn("ID", GetType(Integer)))
                .Add(New DataColumn("Name", GetType(String)))
                .Add(New DataColumn("Score", GetType(Integer)))
                .Add(New DataColumn("Team", GetType(String)))
            End With

            With team1.Rows
                .Add(10, "Bob", 12, "Xi'An")
                .Add(11, "Tommy", 6, "Xi'An")
                .Add(12, "Jaguar", 15, "Xi'An")
                .Add(12, "Lusia", 9, "Xi'An")
            End With

            Dim team2 As New Data.DataTable
            With team2.Columns
                .Add(New DataColumn("ID", GetType(Int32)))
                .Add(New DataColumn("Name", GetType(String)))
                .Add(New DataColumn("Score", GetType(Int32)))
                .Add(New DataColumn("Team", GetType(String)))
            End With

            With team2.Rows
                .Add(2, "Phillip", 9, "BeiJing")
                .Add(3, "Hunter", 10, "BeiJing")
                .Add(4, "Hellen", 8, "BeiJing")
                .Add(5, "Jim", 9, "BeiJing")
            End With

            Dim datasource As New Data.DataSet
            datasource.Tables.Add(team1)
            datasource.Tables.Add(team2)
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
                Return "Template_Scores.xlsx"
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
                Return New String() {"xlsx\Template_Scores.xlsx"}
            End Get
        End Property
    End Class
End Namespace
