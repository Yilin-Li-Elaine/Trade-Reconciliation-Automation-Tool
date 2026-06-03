VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReconciler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dictInternal As Object
Private dictBroker As Object

Private wsInternal As Worksheet
Private wsBroker As Worksheet

Public Sub Init(wsi As Worksheet, wsb As Worksheet)

    Set wsInternal = wsi
    Set wsBroker = wsb

    Set dictInternal = CreateObject("Scripting.Dictionary")
    Set dictBroker = CreateObject("Scripting.Dictionary")

End Sub
Public Sub LoadInternal()

    Dim i As Long
    Dim lastRow As Long
    Dim t As clsTrade
    Dim key As String

    lastRow = wsInternal.Cells(wsInternal.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow

        Set t = New clsTrade

        t.TradeDate = wsInternal.Cells(i, 1).Value
        t.Account = wsInternal.Cells(i, 2).Value
        t.Security = wsInternal.Cells(i, 3).Value
        t.ISIN = wsInternal.Cells(i, 4).Value
        t.Side = wsInternal.Cells(i, 5).Value

        t.Quantity = CDbl(wsInternal.Cells(i, 6).Value)
        t.Price = CDbl(wsInternal.Cells(i, 7).Value)
        t.TradeCcy = wsInternal.Cells(i, 8).Value

        key = BuildKey(t)

        Set dictInternal(key) = t

    Next i

End Sub
Public Sub LoadBroker()

    Dim i As Long
    Dim lastRow As Long
    Dim t As clsTrade
    Dim key As String

    lastRow = wsBroker.Cells(wsBroker.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow

        Set t = New clsTrade

        t.TradeDate = wsBroker.Cells(i, 1).Value
        t.Account = wsBroker.Cells(i, 2).Value
        t.Security = wsBroker.Cells(i, 3).Value
        t.ISIN = wsBroker.Cells(i, 4).Value
        t.Side = wsBroker.Cells(i, 5).Value

        t.Quantity = CDbl(wsBroker.Cells(i, 6).Value)
        t.Price = CDbl(wsBroker.Cells(i, 7).Value)
        t.TradeCcy = wsBroker.Cells(i, 8).Value

        key = BuildKey(t)

        Set dictBroker(key) = t

    Next i

End Sub
Private Function BuildKey(t As clsTrade) As String

    BuildKey = t.TradeDate & "|" & _
               t.Account & "|" & _
               t.ISIN

End Function
Public Sub Reconcile(wsOut As Worksheet)

    Dim k As Variant
    Dim tI As clsTrade
    Dim tB As clsTrade
    Dim r As Long

    wsOut.Cells.Clear

    ' headers
    wsOut.Range("A1:L1").Value = Array( _
        "TradeDate", "Account", "Security", "ISIN", "Side", _
        "Internal Quantity", "Broker Quantity", _
        "Internal Price", "Broker Price", _
        "Internal Currency", "Broker Currency", _
        "Break Type")

    r = 2

    '========================
    ' Internal loop
    '========================
    For Each k In dictInternal.Keys

        If dictBroker.Exists(k) Then

            Set tI = dictInternal(k)
            Set tB = dictBroker(k)

            WriteRow wsOut, r, tI, tB, GetBreakType(tI, tB)
            r = r + 1

        Else

            Set tI = dictInternal(k)

            WriteMissing wsOut, r, tI, "Missing in Broker"
            r = r + 1

        End If

    Next k

    '========================
    ' Broker missing in internal
    '========================
    For Each k In dictBroker.Keys

        If Not dictInternal.Exists(k) Then

            Set tB = dictBroker(k)

            WriteMissingBroker wsOut, r, tB, "Missing in Internal"
            r = r + 1

        End If

    Next k

End Sub
Private Function GetBreakType(a As clsTrade, b As clsTrade) As String

    Dim s As String
    s = ""

    If a.Quantity <> b.Quantity Then s = s & "Quantity Break; "
    If a.Price <> b.Price Then s = s & "Price Break; "
    If a.TradeCcy <> b.TradeCcy Then s = s & "Currency Break; "

    If s = "" Then s = "OK"

    GetBreakType = s

End Function
Private Sub WriteRow(ws As Worksheet, r As Long, a As clsTrade, b As clsTrade, brk As String)

    ws.Cells(r, 1).Value = a.TradeDate
    ws.Cells(r, 2).Value = a.Account
    ws.Cells(r, 3).Value = a.Security
    ws.Cells(r, 4).Value = a.ISIN
    ws.Cells(r, 5).Value = a.Side

    ws.Cells(r, 6).Value = a.Quantity
    ws.Cells(r, 7).Value = b.Quantity

    ws.Cells(r, 8).Value = a.Price
    ws.Cells(r, 9).Value = b.Price

    ws.Cells(r, 10).Value = a.TradeCcy
    ws.Cells(r, 11).Value = b.TradeCcy

    ws.Cells(r, 12).Value = brk

End Sub

Private Sub WriteMissing(ws As Worksheet, r As Long, a As clsTrade, brk As String)

    ws.Cells(r, 1).Value = a.TradeDate
    ws.Cells(r, 2).Value = a.Account
    ws.Cells(r, 3).Value = a.Security
    ws.Cells(r, 4).Value = a.ISIN
    ws.Cells(r, 5).Value = a.Side

    ws.Cells(r, 6).Value = a.Quantity
    ws.Cells(r, 7).Value = 0

    ws.Cells(r, 8).Value = a.Price
    ws.Cells(r, 9).Value = 0

    ws.Cells(r, 10).Value = a.TradeCcy
    ws.Cells(r, 11).Value = ""

    ws.Cells(r, 12).Value = brk

End Sub
Private Sub WriteMissingBroker(ws As Worksheet, r As Long, b As clsTrade, brk As String)

    ws.Cells(r, 1).Value = b.TradeDate
    ws.Cells(r, 2).Value = b.Account
    ws.Cells(r, 3).Value = b.Security
    ws.Cells(r, 4).Value = b.ISIN
    ws.Cells(r, 5).Value = b.Side

    ws.Cells(r, 6).Value = 0
    ws.Cells(r, 7).Value = b.Quantity

    ws.Cells(r, 8).Value = 0
    ws.Cells(r, 9).Value = b.Price

    ws.Cells(r, 10).Value = ""
    ws.Cells(r, 11).Value = b.TradeCcy

    ws.Cells(r, 12).Value = brk

End Sub
