Imports System
Imports System.Collections.Generic

Namespace XLExportExample
    Friend Class InvoiceData
        Public Sub New(ByVal product As String, ByVal qty As Integer, ByVal unitPrice As Double, ByVal discount As Double)
            Me.Product = product
            Me.Qty = qty
            Me.UnitPrice = unitPrice
            Me.Discount = discount
        End Sub

        Private privateProduct As String
        Public Property Product() As String
            Get
                Return privateProduct
            End Get
            Private Set(ByVal value As String)
                privateProduct = value
            End Set
        End Property
        Private privateQty As Integer
        Public Property Qty() As Integer
            Get
                Return privateQty
            End Get
            Private Set(ByVal value As Integer)
                privateQty = value
            End Set
        End Property
        Private privateUnitPrice As Double
        Public Property UnitPrice() As Double
            Get
                Return privateUnitPrice
            End Get
            Private Set(ByVal value As Double)
                privateUnitPrice = value
            End Set
        End Property
        Private privateDiscount As Double
        Public Property Discount() As Double
            Get
                Return privateDiscount
            End Get
            Private Set(ByVal value As Double)
                privateDiscount = value
            End Set
        End Property
    End Class

    Friend Class Invoice

        Private ReadOnly items_Renamed As New List(Of InvoiceData)()

        Public Shared Function CreateSampleInvoice() As Invoice
            Dim result As New Invoice()
            result.InvoiceNum = 100
            result.privateDate = Date.Now
            result.Customer = "Alcorn Mickey"
            result.Company = "Mickeys World of Fun"
            result.Address = "436 1st Ave."
            result.Address2 = "Cleveland, OH, 37288"
            result.Phone = "(203)290-8902"
            result.Fax = "(203)290-8903"
            result.Items.Add(New InvoiceData("Aniseed Syrup", 4, 10, 0))
            result.Items.Add(New InvoiceData("Mishi Kobe Niku", 13, 97, 0.15))
            result.Items.Add(New InvoiceData("Ikura", 12, 31, 0.1))
            result.Items.Add(New InvoiceData("Konbu", 11, 6, 0))
            result.Items.Add(New InvoiceData("Pavlova", 10, 17.45, 0))
            result.Items.Add(New InvoiceData("Boston Crab Meat", 2, 18.4, 0))
            Return result
        End Function

        Private privateInvoiceNum As Integer
        Public Property InvoiceNum() As Integer
            Get
                Return privateInvoiceNum
            End Get
            Private Set(ByVal value As Integer)
                privateInvoiceNum = value
            End Set
        End Property
        Private privateDate As Date
        Public Property [Date]() As Date
            Get
                Return privateDate
            End Get
            Private Set(ByVal value As Date)
                privateDate = value
            End Set
        End Property
        Private privateCustomer As String
        Public Property Customer() As String
            Get
                Return privateCustomer
            End Get
            Private Set(ByVal value As String)
                privateCustomer = value
            End Set
        End Property
        Private privateCompany As String
        Public Property Company() As String
            Get
                Return privateCompany
            End Get
            Private Set(ByVal value As String)
                privateCompany = value
            End Set
        End Property
        Private privateAddress As String
        Public Property Address() As String
            Get
                Return privateAddress
            End Get
            Private Set(ByVal value As String)
                privateAddress = value
            End Set
        End Property
        Private privateAddress2 As String
        Public Property Address2() As String
            Get
                Return privateAddress2
            End Get
            Private Set(ByVal value As String)
                privateAddress2 = value
            End Set
        End Property
        Private privatePhone As String
        Public Property Phone() As String
            Get
                Return privatePhone
            End Get
            Private Set(ByVal value As String)
                privatePhone = value
            End Set
        End Property
        Private privateFax As String
        Public Property Fax() As String
            Get
                Return privateFax
            End Get
            Private Set(ByVal value As String)
                privateFax = value
            End Set
        End Property
        Public ReadOnly Property Items() As List(Of InvoiceData)
            Get
                Return items_Renamed
            End Get
        End Property
    End Class
End Namespace
