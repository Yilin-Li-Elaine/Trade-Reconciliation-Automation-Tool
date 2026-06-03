Attribute VB_Name = "modImport"
Option Explicit

Sub ImportFiles()

    Dim folderPath As String
    Dim internalFile As String
    Dim brokerFile As String

    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    Dim wsInternal As Worksheet
    Dim wsBroker As Worksheet

    '========================
    ' 1. folder path
    '========================
    folderPath = ThisWorkbook.Path & "\"

    '========================
    ' 2. file names
    '========================
    internalFile = Dir(folderPath & "*internal*.csv")
    brokerFile = Dir(folderPath & "*broker*.csv")

    '========================
    ' 3. validation
    '========================
    If internalFile = "" Then
        MsgBox "Internal file not found", vbCritical
        Exit Sub
    End If

    If brokerFile = "" Then
        MsgBox "Broker file not found", vbCritical
        Exit Sub
    End If

    '========================
    ' 4. target sheets
    '========================
    Set wsInternal = ThisWorkbook.Sheets("Internal")
    Set wsBroker = ThisWorkbook.Sheets("Broker")

    ' clear old data
    wsInternal.Cells.Clear
    wsBroker.Cells.Clear

    '========================
    ' 5. import Internal
    '========================
    Set wbSource = Workbooks.Open(folderPath & internalFile)

    Set wsSource = wbSource.Sheets("Internal")

    wsSource.UsedRange.Copy wsInternal.Range("A1")

    wbSource.Close False

    '========================
    ' 6. import Broker
    '========================
    Set wbSource = Workbooks.Open(folderPath & brokerFile)

    Set wsSource = wbSource.Sheets("broker")

    wsSource.UsedRange.Copy wsBroker.Range("A1")

    wbSource.Close False

    '========================
    ' 7. done
    '========================
    MsgBox "Files imported successfully", vbInformation

End Sub
