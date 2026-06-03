Attribute VB_Name = "modMain"
Option Explicit

Sub RunReconciliation()
   
    Call ImportFiles
    Call AlignBrokerColumns
    Call FixBrokerAccount
    
    Dim rec As clsReconciler

    Set rec = New clsReconciler

    rec.Init Sheets("Internal"), Sheets("Broker_output")

    rec.LoadInternal
    rec.LoadBroker

    rec.Reconcile Sheets("Exceptions")
    
    ThisWorkbook.Save
    MsgBox "Reconciliation Complete"
    
    Call BuildDashboard
    Call SendEmail

End Sub
