Attribute VB_Name = "modSoFifa"
Option Explicit

Private oGf As clsGlobalFunc
Private oSf As clsSoFifa

Public Sub SoFifa(ByVal pId As String, ByVal pVersion As String)

Variables:
    Set oSf = New clsSoFifa
    Set oGf = New clsGlobalFunc
    
SearchWebpage:
    Call oSf.Search(pId, pVersion)
    
Finish:
    On Error Resume Next
    Set oSf = Nothing
    Set oGf = Nothing
    Exit Sub
    
ErrHandler:
    Call oGf.DisplayErrMessage("ModSoFifa", "Scrap")
    GoTo Finish

End Sub
