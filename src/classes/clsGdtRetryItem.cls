VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGdtRetryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Direction As String
Public ApiSource As String
Public SellerTaxCode As String
Public TemplateCode As String
Public InvoiceSeries As String
Public InvoiceNumber As String
Public InvoiceDate As Variant
Public Stage As String
Public Endpoint As String
Public ResponseKind As String
Public Attempts As Long
Public TargetRow As Long
Public ActionHeader As String

Public Function UniqueKey() As String
    UniqueKey = BuildGdtErrorKey(Direction, ApiSource, SellerTaxCode, TemplateCode, _
                                 InvoiceSeries, InvoiceNumber, Stage) & "|" & Endpoint
End Function
