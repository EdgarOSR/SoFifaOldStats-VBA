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


Public Sub GameChange()
    Set oCls = New clsForm
    Call oCls.GameChange
    Set oCls = Nothing
End Sub


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


Public Sub TeamChange()
    Set oCls = New clsForm
    Call oCls.TeamChange
    Set oCls = Nothing
End Sub


Public Sub PlayerChange()
    Set oCls = New clsForm
    Call oCls.PlayerChange
    Set oCls = Nothing
End Sub


Public Sub ListboxClick()
    Set oCls = New clsForm
    Call oCls.ListboxClick
    Set oCls = Nothing
End Sub
