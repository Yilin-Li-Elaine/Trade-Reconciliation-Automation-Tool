Attribute VB_Name = "modBrokerData"

Option Explicit

Sub AlignBrokerColumns()

    Dim wsInternal As Worksheet
    Dim wsBroker As Worksheet
    Dim wsOutput As Worksheet

    Dim lastColInternal As Long
    Dim lastRowBroker As Long

    Dim i As Long
    Dim brokerCol As Long

    Dim headerName As String

    Set wsInternal = Sheets("Internal")
    Set wsBroker = Sheets("Broker")

    ' output sheet
    Set wsOutput = Sheets("Broker_output")
    wsOutput.Cells.Clear

    '========================
    ' Internal standard header
    '========================
    lastColInternal = wsInternal.Cells(1, wsInternal.Columns.Count).End(xlToLeft).Column

    lastRowBroker = wsBroker.Cells(wsBroker.Rows.Count, 1).End(xlUp).row

    '========================
    ' Internal header order
    '========================
    For i = 1 To lastColInternal

        headerName = Trim(UCase(wsInternal.Cells(1, i).Value))

        ' find Broker col
        brokerCol = FindColumn(wsBroker, headerName)

        If brokerCol = 0 Then

            MsgBox "Missing column in Broker: " & headerName, vbCritical
            Exit Sub

        End If

        ' copy col
        wsBroker.Range( _
            wsBroker.Cells(1, brokerCol), _
            wsBroker.Cells(lastRowBroker, brokerCol) _
        ).Copy wsOutput.Cells(1, i)

    Next i

    
End Sub
Sub FixBrokerAccount()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim colAccount As Long
    Dim i As Long

    Set ws = Sheets("Broker_output")

    ' find account col
    colAccount = FindColumn(ws, "ACCOUNT")

    If colAccount = 0 Then

        MsgBox "ACCOUNT column not found", vbCritical
        Exit Sub

    End If

    ' find last row
    lastRow = ws.Cells(ws.Rows.Count, colAccount).End(xlUp).row

    ' loop
    For i = 2 To lastRow

        ws.Cells(i, colAccount).Value = _
            NormalizeAccount(ws.Cells(i, colAccount).Value)

    Next i

End Sub
