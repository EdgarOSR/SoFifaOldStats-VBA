Attribute VB_Name = "modForm"
Option Explicit
Option Private Module


Private oCls As clsForm


Public Function FormClose() As Integer
    Set oCls = New clsForm
    Let FormClose = oCls.FormClose
    Set oCls = Nothing
End Function


Public Sub BtnCopyClick()
    Set oCls = New clsForm
    Call oCls.BtnCopyClick
    Set oCls = Nothing
End Sub


Public Sub BtnClearClick()
    Set oCls = New clsForm
    Call oCls.BtnClearClick
    Set oCls = Nothing
End Sub


Public Sub BtnSearchClick()
    Set oCls = New clsForm
    Call oCls.BtnSearchClick
    Set oCls = Nothing
End Sub


Public Sub FormActivate()
    Set oCls = New clsForm
    Call oCls.FormActivate
    Set oCls = Nothing
End Sub


'Public Sub SearchClubsList()
'    Set oCls = New clsForm
'    Call oCls.SearchClubsList
'    Set oCls = Nothing
'End Sub
'
'
'Public Sub SearchPlayersList()
'    Set oCls = New clsForm
'    Call oCls.SearchPlayersList
'    Set oCls = Nothing
'End Sub

Public Sub VersionChange()
    Set oCls = New clsForm
    Call oCls.VersionChange
    Set oCls = Nothing
End Sub


Public Sub CountryChange()
    Set oCls = New clsForm
    Call oCls.CountryChange
    Set oCls = Nothing
End Sub


Public Sub ClubChange()
    Set oCls = New clsForm
    Call oCls.ClubChange
    Set oCls = Nothing
End Sub


Public Sub PlayerChange()
    Set oCls = New clsForm
    Call oCls.PlayerChange
    Set oCls = Nothing
End Sub
