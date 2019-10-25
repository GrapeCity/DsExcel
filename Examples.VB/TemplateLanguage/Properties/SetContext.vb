Imports System.Data

Namespace Templates.Properties
    Public Class SetContext
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file Template_Score.xlsx from resource
            Dim templateFile = GetResourceStream("xlsx\Template_SetContext.xlsx")
            workbook.Open(templateFile)

#Region "Init Data"
            Dim datasource = New DataTable
            With datasource.Columns
                .Add(New DataColumn("ID", GetType(Integer)))
                .Add(New DataColumn("Name", GetType(String)))
                .Add(New DataColumn("Score", GetType(Integer)))
                .Add(New DataColumn("Team", GetType(String)))
            End With

            With datasource.Rows
                .Add(10, "Bob", 12, "Xi'An")
                .Add(11, "Tommy", 6, "Xi'An")
                .Add(12, "Jaguar", 15, "Xi'An")
                .Add(2, "Phillip", 9, "BeiJing")
                .Add(3, "Hunter", 10, "BeiJing")
                .Add(4, "Hellen", 8, "BeiJing")
                .Add(5, "Jim", 9, "BeiJing")
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
                Return "Template_SetContext.xlsx"
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
                Return New String() {"xlsx\Template_SetContext.xlsx"}
            End Get
        End Property
    End Class
End Namespace
