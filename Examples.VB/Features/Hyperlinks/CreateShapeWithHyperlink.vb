Namespace Features.Hyperlinks
    Public Class CreateShapeWithHyperlink
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Excel.Workbook)
            Dim worksheet As IWorksheet = workbook.Worksheets(0)

            ' Add shapes
            Dim shape1 As IShape = worksheet.Shapes.AddShape(AutoShapeType.Rectangle, 10, 0, 100, 100)
            shape1.TextFrame.TextRange.Add("Go to google web site.")
            Dim shape2 As IShape = worksheet.Shapes.AddShape(AutoShapeType.RightArrow, 10, 120, 100, 100)
            shape2.TextFrame.TextRange.Add("Go to sheet1 C3:E4")
            Dim shape3 As IShape = worksheet.Shapes.AddShape(AutoShapeType.Oval, 10, 240, 100, 100)
            shape3.TextFrame.TextRange.Add("Send an email to sales")
            Dim shape4 As IShape = worksheet.Shapes.AddShape(AutoShapeType.LeftArrow, 10, 360, 100, 100)
            shape4.TextFrame.TextRange.Add("Link to external.xlsx file.")

            With worksheet.Hyperlinks
                'add a hyperlink link to web page.
                .Add(shape1, "https://www.google.com/", , "open google web site.", "Google")

                'add a hyperlink link to a range in this document.
                .Add(shape2, Nothing, "Sheet1!$C$3:$E$4", "Go to sheet1 C3:E4")

                'add a hyperlink link to email address.
                .Add(shape3, "mailto:us.sales@grapecity.com", , "Send an email to sales", "Send an email to sales")

                'add a hyperlink link to external file.
                'change the path to real picture file path.
                .Add(shape4, address:="external.xlsx", screenTip:="link to external.xlsx file.",
                     textToDisplay:="External.xlsx")
            End With
        End Sub

        Public Overrides ReadOnly Property ShowViewer As Boolean
            Get
                Return False
            End Get
        End Property
    End Class
End Namespace
