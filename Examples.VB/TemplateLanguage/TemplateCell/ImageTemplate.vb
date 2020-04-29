Namespace Templates.TemplateCell
    Public Class ImageTemplate
        Inherits ExampleBase

        Public Overrides Sub Execute(workbook As Workbook)
            'Load template file from resource
            Dim templateFile = GetResourceStream("xlsx\Template_ImageTemplate.xlsx")
            workbook.Open(templateFile)

#Region "Define custom classes"
            'Public Class BikeInfo
            '	Public BikeType As String
            '	Public BikeSeries As List(Of BikeSeries)
            'End Class

            'Public Class BikeSeries
            '	Public Name As String
            '	Public Description As String
            '	Public BikeImage() As Byte
            '	Public Items As List(Of Bike)
            'End Class

            'Public Class Bike
            '	Public ProductNo As String
            '	Public ProductName As String
            '	Public Color As String
            '	Public Size As Integer
            '	Public Weight As Double
            '	Public Dealer As Double
            '	Public ListPrice As Double
            'End Class		
#End Region

#Region "Init Data"
            Dim image1 As Byte()
            Using imageStream1 = GetResourceStream("image\Mountain-100.jpg")
                ReDim image1(CInt(imageStream1.Length) - 1)
                imageStream1.Read(image1, 0, image1.Length)
            End Using
            Dim image2 As Byte()
            Using imageStream2 = GetResourceStream("image\Mountain-200.jpg")
                ReDim image2(CInt(imageStream2.Length) - 1)
                imageStream2.Read(image2, 0, image2.Length)
            End Using
            Dim image3 As Byte()
            Using imageStream3 = GetResourceStream("image\Mountain-300.jpg")
                ReDim image3(CInt(imageStream3.Length) - 1)
                imageStream3.Read(image3, 0, image3.Length)
            End Using
            Dim image4 As Byte()
            Using imageStream4 = GetResourceStream("image\Mountain-400-W.jpg")
                ReDim image4(CInt(imageStream4.Length) - 1)
                imageStream4.Read(image4, 0, image4.Length)
            End Using
            Dim image5 As Byte()
            Using imageStream5 = GetResourceStream("image\Mountain-500.jpg")
                ReDim image5(CInt(imageStream5.Length) - 1)
                imageStream5.Read(image5, 0, image5.Length)
            End Using
            Dim image6 As Byte()
            Using imageStream6 = GetResourceStream("image\Road-150.jpg")
                ReDim image6(CInt(imageStream6.Length) - 1)
                imageStream6.Read(image6, 0, image6.Length)
            End Using
            Dim image7 As Byte()
            Using imageStream7 = GetResourceStream("image\Road-350-W.jpg")
                ReDim image7(CInt(imageStream7.Length) - 1)
                imageStream7.Read(image7, 0, image7.Length)
            End Using
            Dim image8 As Byte()
            Using imageStream8 = GetResourceStream("image\Touring-1000.jpg")
                ReDim image8(CInt(imageStream8.Length) - 1)
                imageStream8.Read(image8, 0, image8.Length)
            End Using
            Dim image9 As Byte()
            Using imageStream9 = GetResourceStream("image\Touring-2000.jpg")
                ReDim image9(CInt(imageStream9.Length) - 1)
                imageStream9.Read(image9, 0, image9.Length)
            End Using

            Dim datasource As New List(Of BikeInfo)

            Dim bike1 As New BikeInfo
            datasource.Add(bike1)
            bike1.BikeType = "Mountain Bikes"
            bike1.BikeSeries = New List(Of BikeSeries)

            Dim bs1 As New BikeSeries
            bike1.BikeSeries.Add(bs1)
            bs1.Name = "Mountain-100"
            bs1.BikeImage = image1
            bs1.Description = "Top-of-the-line competition mountain bike. Performance-enhancing options include the innovative HL Frame, super-smooth front suspension, and traction for all terrain."
            bs1.Items = New List(Of Bike)
            Dim bItem1 = New Bike With {
                .ProductNo = "BK-M82S-38",
                .ProductName = "Mountain-100 Silver, 38",
                .Color = "Silver",
                .Size = 38,
                .Weight = 20.35,
                .Dealer = 1912.1544,
                .ListPrice = 3399.99
            }
            bs1.Items.Add(bItem1)
            Dim bItem2 = New Bike With {
                .ProductNo = "BK-M82B-38",
                .ProductName = "Mountain-100 Black, 38",
                .Color = "Black",
                .Size = 38,
                .Weight = 20.35,
                .Dealer = 1898.0944,
                .ListPrice = 3374.99
            }
            bs1.Items.Add(bItem2)

            Dim bs2 As New BikeSeries
            bike1.BikeSeries.Add(bs2)
            bs2.Name = "Mountain-200"
            bs2.BikeImage = image2
            bs2.Description = "Serious back-country riding. Perfect for all levels of competition. Uses the same HL Frame as the Mountain-100."
            bs2.Items = New List(Of Bike)
            Dim bItem3 = New Bike With {
                .ProductNo = "BK-M68S-42",
                .ProductName = "Mountain-200 Silver, 42",
                .Color = "Silver",
                .Size = 42,
                .Weight = 23.77,
                .Dealer = 1265.6195,
                .ListPrice = 2319.99
            }
            bs2.Items.Add(bItem3)
            Dim bItem4 = New Bike With {
                .ProductNo = "BK-M68B-38",
                .ProductName = "Mountain-200 Black, 38",
                .Color = "Black",
                .Size = 38,
                .Weight = 23.35,
                .Dealer = 1251.9813,
                .ListPrice = 2294.99
            }
            bs2.Items.Add(bItem4)

            Dim bs3 As New BikeSeries
            bike1.BikeSeries.Add(bs3)
            bs3.Name = "Mountain-300"
            bs3.BikeImage = image3
            bs3.Description = "For true trail addicts.  An extremely durable bike that will go anywhere and keep you in control on challenging terrain - without breaking your budget."
            bs3.Items = New List(Of Bike)
            Dim bItem5 = New Bike With {
                .ProductNo = "BK-M47B-38",
                .ProductName = "Mountain-300 Black, 38",
                .Color = "Black",
                .Size = 38,
                .Weight = 25.35,
                .Dealer = 598.4354,
                .ListPrice = 1079.99
            }
            bs3.Items.Add(bItem5)
            Dim bItem6 = New Bike With {
                .ProductNo = "BK-M47B-40",
                .ProductName = "Mountain-300 Black, 40",
                .Color = "Black",
                .Size = 40,
                .Weight = 25.77,
                .Dealer = 598.4354,
                .ListPrice = 1079.99
            }
            bs3.Items.Add(bItem6)

            Dim bs4 As New BikeSeries
            bike1.BikeSeries.Add(bs4)
            bs4.Name = "Mountain-400-W"
            bs4.BikeImage = image4
            bs4.Description = "This bike delivers a high-level of performance on a budget. It is responsive and maneuverable, and offers peace-of-mind when you decide to go off-road."
            bs4.Items = New List(Of Bike)
            Dim bItem7 = New Bike With {
                .ProductNo = "BKBK-M38S-38",
                .ProductName = "Mountain-400-W Silver, 38",
                .Color = "Silver",
                .Size = 38,
                .Weight = 26.35,
                .Dealer = 419.7784,
                .ListPrice = 769.49
            }
            bs4.Items.Add(bItem7)
            Dim bItem8 = New Bike With {
                .ProductNo = "BK-M38S-40",
                .ProductName = "Mountain-400-W Silver, 40",
                .Color = "Silver",
                .Size = 40,
                .Weight = 26.77,
                .Dealer = 419.7784,
                .ListPrice = 769.49
            }
            bs4.Items.Add(bItem8)

            Dim bs5 As New BikeSeries
            bike1.BikeSeries.Add(bs5)
            bs5.Name = "Mountain-500"
            bs5.BikeImage = image5
            bs5.Description = "Suitable for any type of riding, on or off-road. Fits any budget. Smooth-shifting with a comfortable ride."
            bs5.Items = New List(Of Bike)
            Dim bItem9 = New Bike With {
                .ProductNo = "BK-M18S-40",
                .ProductName = "Mountain-500 Silver, 40",
                .Color = "Silver",
                .Size = 40,
                .Weight = 27.35,
                .Dealer = 308.2179,
                .ListPrice = 564.99
            }
            bs5.Items.Add(bItem9)
            Dim bItem10 = New Bike With {
                .ProductNo = "BK-M18B-40",
                .ProductName = "Mountain-500 Black, 40",
                .Color = "Black",
                .Size = 40,
                .Weight = 27.35,
                .Dealer = 294.5797,
                .ListPrice = 539.99
            }
            bs5.Items.Add(bItem10)


            Dim bike2 = New BikeInfo()
            datasource.Add(bike2)
            bike2.BikeType = "Road Bikes"
            bike2.BikeSeries = New List(Of BikeSeries)()

            Dim bs6 As New BikeSeries
            bike2.BikeSeries.Add(bs6)
            bs6.Name = "Road-150"
            bs6.BikeImage = image6
            bs6.Description = "This bike is ridden by race winners. Developed with the Adventure Works Cycles professional race team, it has a extremely light heat-treated aluminum frame, and steering that allows precision control."
            bs6.Items = New List(Of Bike)
            Dim bItem11 = New Bike With {
                .ProductNo = "BK-R93R-62",
                .ProductName = "Road-150 Red, 62",
                .Color = "Red",
                .Size = 62,
                .Weight = 15,
                .Dealer = 2171.2942,
                .ListPrice = 3578.27
            }
            bs6.Items.Add(bItem11)
            Dim bItem12 = New Bike With {
                .ProductNo = "BK-R93R-44",
                .ProductName = "Road-150 Red, 44",
                .Color = "Red",
                .Size = 44,
                .Weight = 13.77,
                .Dealer = 2171.2942,
                .ListPrice = 3578.27
            }
            bs6.Items.Add(bItem12)

            Dim bs7 As New BikeSeries
            bike2.BikeSeries.Add(bs7)
            bs7.Name = "Road-350-W"
            bs7.BikeImage = image7
            bs7.Description = "Cross-train, race, or just socialize on a sleek, aerodynamic bike designed for a woman.  Advanced seat technology provides comfort all day."
            bs7.Items = New List(Of Bike)
            Dim bItem13 = New Bike With {
                .ProductNo = "BK-R79Y-40",
                .ProductName = "Road-350-W Yellow, 40",
                .Color = "Yellow",
                .Size = 40,
                .Weight = 15.35,
                .Dealer = 1082.51,
                .ListPrice = 1700.99
            }
            bs7.Items.Add(bItem13)
            Dim bItem14 = New Bike With {
                .ProductNo = "BK-R79Y-42",
                .ProductName = "Road-350-W Yellow, 42",
                .Color = "Yellow",
                .Size = 42,
                .Weight = 15.77,
                .Dealer = 1082.51,
                .ListPrice = 1700.99
            }
            bs7.Items.Add(bItem14)


            Dim bike3 = New BikeInfo()
            datasource.Add(bike3)
            bike3.BikeType = "Touring Bikes"
            bike3.BikeSeries = New List(Of BikeSeries)()

            Dim bs8 As New BikeSeries
            bike3.BikeSeries.Add(bs8)
            bs8.Name = "Touring-1000"
            bs8.BikeImage = image8
            bs8.Description = "Travel in style and comfort. Designed for maximum comfort and safety. Wide gear range takes on all hills. High-tech aluminum alloy construction provides durability without added weight."
            bs8.Items = New List(Of Bike)
            Dim bItem15 = New Bike With {
                .ProductNo = "BK-T79Y-46",
                .ProductName = "Touring-1000 Yellow, 46",
                .Color = "Yellow",
                .Size = 46,
                .Weight = 25.13,
                .Dealer = 1481.9379,
                .ListPrice = 2384.07
            }
            bs8.Items.Add(bItem15)
            Dim bItem16 = New Bike With {
                .ProductNo = "BK-T79U-46",
                .ProductName = "Touring-1000 Blue, 46",
                .Color = "Blue",
                .Size = 46,
                .Weight = 25.13,
                .Dealer = 1481.9379,
                .ListPrice = 2384.07
            }
            bs8.Items.Add(bItem16)

            Dim bs9 As New BikeSeries
            bike3.BikeSeries.Add(bs9)
            bs9.Name = "Touring-2000"
            bs9.BikeImage = image9
            bs9.Description = "The plush custom saddle keeps you riding all day,  and there's plenty of space to add panniers and bike bags to the newly-redesigned carrier.  This bike has stability when fully-loaded."
            bs9.Items = New List(Of Bike)
            Dim bItem17 = New Bike With {
                .ProductNo = "BK-T44U-60",
                .ProductName = "Touring-2000 Blue, 60",
                .Color = "Blue",
                .Size = 60,
                .Weight = 27.9,
                .Dealer = 755.1508,
                .ListPrice = 1214.85
            }
            bs9.Items.Add(bItem17)
#End Region

            'Init template global settings
            workbook.Names.Add("TemplateOptions.KeepLineSize", "true")

            'Add data source
            workbook.AddDataSource("ds", datasource)
            'Invoke to process the template
            workbook.ProcessTemplate()
        End Sub

        Public Overrides ReadOnly Property TemplateName() As String
            Get
                Return "Template_ImageTemplate.xlsx"
            End Get
        End Property

        Public Overrides ReadOnly Property HasTemplate() As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides ReadOnly Property CanDownloadZip() As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides ReadOnly Property UsedResources() As String()
            Get
                Return New String() {"xlsx\Template_ImageTemplate.xlsx", "image\Mountain-100.jpg", "image\Mountain-200.jpg", "image\Mountain-300.jpg", "image\Mountain-400-W.jpg", "image\Mountain-500.jpg", "image\Road-150.jpg", "image\Road-350-W.jpg", "image\Touring-1000.jpg", "image\Touring-2000.jpg"}
            End Get
        End Property

        Public Overrides ReadOnly Property Refs() As String()
            Get
                Return New String() {"BikeInfo", "BikeSeries", "Bike"}
            End Get
        End Property

        Public Overrides ReadOnly Property IsNew() As Boolean
            Get
                Return True
            End Get
        End Property
    End Class

    Public Class BikeInfo
        Public BikeType As String
        Public BikeSeries As List(Of BikeSeries)
    End Class

    Public Class BikeSeries
        Public Name As String
        Public Description As String
        Public BikeImage() As Byte
        Public Items As List(Of Bike)
    End Class

    Public Class Bike
        Public ProductNo As String
        Public ProductName As String
        Public Color As String
        Public Size As Integer
        Public Weight As Double
        Public Dealer As Double
        Public ListPrice As Double
    End Class
End Namespace
