Attribute VB_Name = "modEmail"
Sub SendEmail()

    Dim OutlookApp As Object
    Dim OutlookMail As Object

    '========================
    ' 1. Write Email
    '========================
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    With OutlookMail

        .To = "xxx@email.com"

        .Subject = "Daily Reconciliation Completed"

        .Body = "Please find attached reconciliation report."
        
        '========================
        ' 2. Attachment
        '========================

        .Attachments.Add ThisWorkbook.FullName
        
        .Display
        .Send

    End With

    MsgBox "Email sent successfully", vbInformation
    
    '========================
    ' 3. Cleanup
    '========================
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
