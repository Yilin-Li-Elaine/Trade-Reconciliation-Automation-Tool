Attribute VB_Name = "ModDashboard"
Public Sub BuildDashboard()

    Dim wsEx As Worksheet
    Dim wsDash As Worksheet

    Dim lastRow As Long
    Dim i As Long

    Dim totalTrades As Long
    Dim okCount As Long
    Dim qtyBreak As Long
    Dim priceBreak As Long
    Dim ccyBreak As Long
    Dim missBroker As Long
    Dim missInternal As Long
    
    Const COL_BREAKTYPE As Long = 12
    
    Dim breakType As String
    Dim matchRate As Double


    Set wsEx = Sheets("Exceptions")

    '========================
    ' Create Dashboard Sheet
    '========================
    On Error Resume Next
    Set wsDash = Sheets("Dashboard")
    On Error GoTo 0

    If wsDash Is Nothing Then
        Set wsDash = Sheets.Add
        wsDash.Name = "Dashboard"
    End If

    wsDash.Cells.Clear

    '========================
    ' Count Breaks
    '========================

    lastRow = wsEx.Cells(wsEx.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRow

        breakType = Trim(wsEx.Cells(i, COL_BREAKTYPE).Value)

        totalTrades = totalTrades + 1
        
        Select Case breakType

            Case "OK"
                okCount = okCount + 1

            Case "Missing in Broker"
                missBroker = missBroker + 1

            Case "Missing in Internal"
                missInternal = missInternal + 1

        End Select
        
        
        If InStr(1, breakType, "Quantity Break") > 0 Then

            qtyBreak = qtyBreak + 1
        
        End If
        
        If InStr(1, breakType, "Price Break") > 0 Then
        
            priceBreak = priceBreak + 1
        
        End If
        
        If InStr(1, breakType, "Currency Break") > 0 Then
        
            ccyBreak = ccyBreak + 1
        
        End If

    Next i

    If totalTrades > 0 Then
        matchRate = okCount / totalTrades
    End If

    '========================
    ' Dashboard Title
    '========================

    With wsDash.Range("A1:H1")

        .Merge
        .Value = "TRADE RECONCILIATION DASHBOARD"

        .Font.Size = 18
        .Font.Bold = True

    End With

    '========================
    ' KPI Section
    '========================

    wsDash.Range("A3").Value = "Total Trades"
    wsDash.Range("B3").Value = totalTrades

    wsDash.Range("A4").Value = "Matched Trades"
    wsDash.Range("B4").Value = okCount

    wsDash.Range("A5").Value = "Match Rate"
    wsDash.Range("B5").Value = matchRate
    wsDash.Range("B5").NumberFormat = "0.00%"

    If matchRate >= 0.95 Then
        wsDash.Range("B5").Interior.Color = vbGreen
    Else
        wsDash.Range("B5").Interior.Color = vbRed
    End If

    '========================
    ' Break Summary Table
    '========================

    wsDash.Range("D3").Value = "Break Type"
    wsDash.Range("E3").Value = "Count"

    wsDash.Range("D4").Value = "Quantity Break"
    wsDash.Range("E4").Value = qtyBreak

    wsDash.Range("D5").Value = "Price Break"
    wsDash.Range("E5").Value = priceBreak

    wsDash.Range("D6").Value = "Currency Break"
    wsDash.Range("E6").Value = ccyBreak

    wsDash.Range("D7").Value = "Missing in Broker"
    wsDash.Range("E7").Value = missBroker

    wsDash.Range("D8").Value = "Missing in Internal"
    wsDash.Range("E8").Value = missInternal

    '========================
    ' Data for Pie Chart
    '========================

    wsDash.Range("G3").Value = "Category"
    wsDash.Range("H3").Value = "Count"

    wsDash.Range("G4").Value = "Matched"
    wsDash.Range("H4").Value = okCount

    wsDash.Range("G5").Value = "Exceptions"
    wsDash.Range("H5").Value = totalTrades - okCount

    '========================
    ' Delete Existing Charts
    '========================

    Dim cht As ChartObject

    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht

    '========================
    ' Pie Chart
    '========================

    Dim pieChart As ChartObject

    Set pieChart = wsDash.ChartObjects.Add(20, 180, 300, 220)

    With pieChart.Chart

        .SetSourceData wsDash.Range("G3:H5")
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Matched vs Exceptions"

    End With

    '========================
    ' Break Distribution Chart
    '========================

    Dim barChart As ChartObject

    Set barChart = wsDash.ChartObjects.Add(20, 450, 450, 250)

    With barChart.Chart

        .SetSourceData wsDash.Range("D3:E8")
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Break Distribution"

    End With

    '========================
    ' Refresh Time
    '========================

    wsDash.Range("A10").Value = "Last Refresh"

    wsDash.Range("B10").Value = _
        Format(Now, "yyyy-mm-dd hh:mm:ss")

    '========================
    ' Formatting
    '========================

    wsDash.Columns("A:H").AutoFit

    MsgBox "Dashboard Updated"

End Sub

