Imports System.Data

Namespace Features.DataBinding.SheetBinding
    Public Class BindDataTable
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
#Region "Init data"
            Dim teamInfo As New DataTable
            With teamInfo.Columns
                .Add(New DataColumn("ID", GetType(Integer)))
                .Add(New DataColumn("Name", GetType(String)))
                .Add(New DataColumn("Score", GetType(Integer)))
                .Add(New DataColumn("Team", GetType(String)))
            End With

            With teamInfo.Rows
                .Add(10, "Bob", 12, "Xi'An")
                .Add(11, "Tommy", 6, "Xi'An")
                .Add(12, "Jaguar", 15, "Xi'An")
                .Add(12, "Lusia", 9, "Xi'An")
            End With
#End Region

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Set AutoGenerateColumns as false
            worksheet.AutoGenerateColumns = False

            ' Bind columns manually.
            worksheet.Range("A:A").EntireColumn.BindingPath = "ID"
            worksheet.Range("B:B").EntireColumn.BindingPath = "Name"
            worksheet.Range("C:C").EntireColumn.BindingPath = "Score"
            worksheet.Range("D:D").EntireColumn.BindingPath = "Team"

            ' Set data source
            worksheet.DataSource = teamInfo
        End Sub
    End Class
End Namespace
