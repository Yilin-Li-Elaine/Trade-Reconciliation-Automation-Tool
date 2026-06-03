Attribute VB_Name = "modUtility"
Option Explicit

Sub ClearOldData()

    Sheets("Internal").Cells.ClearContents
    Sheets("Broker").Cells.ClearContents
    Sheets("Broker_output").Cells.ClearContents
    Sheets("Exceptions").Cells.ClearContents

End Sub

Function FindColumn(ws As Worksheet, headerName As String) As Long

    Dim j As Long
    Dim lastCol As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For j = 1 To lastCol

        If Trim(UCase(ws.Cells(1, j).Value)) = headerName Then

            FindColumn = j
            Exit Function

        End If

    Next j

    FindColumn = 0

End Function


Function NormalizeAccount(Account As String) As String

    Account = Trim(UCase(Account))

    ' Broker account normalize
    If Right(Account, 2) = "PF" Then

        Account = "PF" & Left(Account, Len(Account) - 2)

    End If

    NormalizeAccount = Account

End Function

Function BuildKey( _
    TradeDate As String, _
    Account As String, _
    ISIN As String) As String
    ' Return to a string e.g.2026-05-27|PF1001|US0378331005
    BuildKey = Trim(TradeDate) & "|" & _
               Trim(Account) & "|" & _
               Trim(ISIN)

End Function
