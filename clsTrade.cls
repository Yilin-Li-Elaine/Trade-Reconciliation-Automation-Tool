VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TradeDate As String
Public Account As String
Public Security As String
Public ISIN As String
Public Side As String

Public Quantity As Double
Public Price As Double
Public TradeCcy As String
