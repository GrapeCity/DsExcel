Imports System.Data

Namespace Features.DataBinding.TableBinding
    Public Class BindDataSet
        Inherits ExampleBase
        Public Overrides Sub Execute(workbook As Excel.Workbook)
#Region "Init data"
            ' DataSet
            Dim team1 As New DataTable("T1")
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

            Dim team2 As New DataTable("T2")
            With team2.Columns
                .Add(New DataColumn("ID", GetType(Integer)))
                .Add(New DataColumn("Name", GetType(String)))
                .Add(New DataColumn("Score", GetType(Integer)))
                .Add(New DataColumn("Team", GetType(String)))
            End With

            With team2.Rows
                .Add(2, "Phillip", 9, "BeiJing")
                .Add(3, "Hunter", 10, "BeiJing")
                .Add(4, "Hellen", 8, "BeiJing")
                .Add(5, "Jim", 9, "BeiJing")
            End With

            Dim datasource As New DataSet
            datasource.Tables.Add(team1)
            datasource.Tables.Add(team2)
#End Region

            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Add tables
            Dim table As ITable = worksheet.Tables.Add(worksheet.Range("B2:E6"), True)
            Dim table2 As ITable = worksheet.Tables.Add(worksheet.Range("G2:J6"), True)

            ' Set not to auto generate table columns
            table.AutoGenerateColumns = False
            table2.AutoGenerateColumns = False

            ' Set table binding path
            table.BindingPath = "T1"
            table2.BindingPath = "T2"

            ' Set table column data field
            table.Columns(0).DataField = "ID"
            table.Columns(1).DataField = "Name"
            table.Columns(2).DataField = "Score"
            table.Columns(3).DataField = "Team"

            table2.Columns(0).DataField = "ID"
            table2.Columns(1).DataField = "Name"
            table2.Columns(2).DataField = "Score"
            table2.Columns(3).DataField = "Team"

            ' Set DataSet as datasource
            worksheet.DataSource = datasource
        End Sub
    End Class
End Namespace
