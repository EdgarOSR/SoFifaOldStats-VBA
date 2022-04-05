VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoFifaLeagues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const clubBaseUrl As String = "https://sofifa.com/team/{club}?r={version}&set=true"
Private Const compBaseUrl As String = "https://sofifa.com/teams?{country}&r={version}&set=true"
Private oDicClubs As Scripting.Dictionary
Private oDicPlayers As Scripting.Dictionary
Private oForm As frSoFifa


Public Sub SearchClubsList()

    Dim oDoc As MSHTML.HTMLDocument
    Dim vUrl As String
    
    Set oForm = frSoFifa
    Set oDoc = New MSHTML.HTMLDocument
    
    Let vUrl = compBaseUrl
    Let vUrl = Replace(vUrl, "{version}", oForm.txVersion.Text)
    Let vUrl = Replace(vUrl, "{country}", oForm.txCountryHidden.Text)

    Set oDoc = GetDocHtml(vUrl)
    Set oDicClubs = GetClubsList(oDoc)
    Call LoadClubs(oDicClubs)
    Call LoadClubsHidden(oDicClubs)

Finish:
    On Error Resume Next
    Set oDoc = Nothing
    Set oDicClubs = Nothing
    Set oForm = Nothing
    Exit Sub

ErrHandler:
    Call MsgBox("Error", vbExclamation)
    GoTo Finish

End Sub


Public Sub SearchPlayersList()

    Dim oDoc As MSHTML.HTMLDocument
    Dim vUrl As String

    Set oForm = frSoFifa
    Set oDoc = New MSHTML.HTMLDocument
    
    Let vUrl = clubBaseUrl
    Let vUrl = Replace(vUrl, "{version}", oForm.txVersion.Text)
    Let vUrl = Replace(vUrl, "{club}", oForm.txClubHidden.Text)

    Set oDoc = GetDocHtml(vUrl)
    Set oDicPlayers = GetPlayersList(oDoc)
    Call LoadPlayers(oDicPlayers)
    Call LoadPlayersHidden(oDicPlayers)

Finish:
    On Error Resume Next
    Set oDoc = Nothing
    Set oDicPlayers = Nothing
    Set oForm = Nothing
    Exit Sub

ErrHandler:
    Call MsgBox("Error", vbExclamation)
    GoTo Finish

End Sub


Private Function GetDocHtml(ByVal pUrl As String) As MSHTML.HTMLDocument

    Dim oIe As MSXML2.XMLHTTP60
    Dim oDoc As MSHTML.HTMLDocument
    Set oIe = New MSXML2.XMLHTTP60
    Set oDoc = New MSHTML.HTMLDocument

    On Error GoTo ErrHandler

    With oIe
        Call .Open("GET", pUrl, False)
        Call .send
        Do While (.readyState <> 4): DoEvents: Loop
        If (.Status <> 200) Then GoTo ErrHandler
    End With

    Let oDoc.body.innerHTML = oIe.responseText

    Set GetDocHtml = oDoc

    On Error GoTo 0

Finish:
    On Error Resume Next
    Set oIe = Nothing
    Set oDoc = Nothing
    Exit Function

ErrHandler:
    Call MsgBox(oIe.statusText, vbExclamation, oIe.Status)
    GoTo Finish

End Function


Private Function GetClubsList(ByVal pDoc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim oDic As Scripting.Dictionary
    Dim vMax As Integer
    Dim vCount As Integer
    Dim vAux As String
    Dim vId As Double
    Set oDic = New Scripting.Dictionary

    Let vMax = pDoc.querySelectorAll(".card .col-name-wide a").length - 2

    For vCount = 0 To vMax Step 2
        With pDoc.querySelectorAll(".card .col-name-wide a")
            Let vId = 0
            Let vId = Split(.Item(vCount).href, "/")(2)
            Let vAux = vbNullString
            Let vAux = .Item(vCount).innerText
        End With
        Call oDic.Add(vId, vAux)
    Next vCount

    Set GetClubsList = oDic

    Set oDic = Nothing

End Function


Private Function GetPlayersList(ByVal pDoc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim oDic As Scripting.Dictionary
    Dim vMax As Integer
    Dim vCount As Integer
    Dim vAux As String
    Dim vId As Double

    Set oDic = New Scripting.Dictionary
    Let vMax = pDoc.querySelectorAll("tbody > tr > td.col-name").length - 1

    For vCount = 0 To vMax Step 2
        With pDoc.querySelectorAll("tbody > tr > td.col-name")
            Let vAux = vbNullString
            Let vAux = vAux & Trim(Split(Split(.Item(vCount + 1).innerText, vbCrLf)(0))(1))
            Let vAux = vAux & Space(1) & Trim(Split(.Item(vCount).innerText, vbCrLf)(1))
            Let vAux = vAux & Space(1) & Trim(Split(.Item(vCount).innerText, vbCrLf)(0))
            Let vId = 0
            Let vId = Val(Split(.Item(vCount).firstChild.href, "/")(2))
        End With
        Call oDic.Add(vId, vAux)
    Next vCount

    Set GetPlayersList = oDic

    Set oDic = Nothing

End Function


Private Sub LoadClubs(ByVal pDic As Scripting.Dictionary)
    
    Dim vCount As Integer

    Set oForm = frSoFifa

    If (pDic.Count = 0) Then
        Let oForm.txClub.ListIndex = -1
        GoTo Finish
    End If
    
    With oForm.txClub
        Call .Clear
        For vCount = 0 To pDic.Count
            If (vCount = pDic.Count) Then Exit For
            Call .AddItem(pDic.Items(vCount), vCount)
        Next vCount
        Let .ListIndex = 0
    End With

Finish:
    Set oForm = Nothing
    Set pDic = Nothing

End Sub


Private Sub LoadClubsHidden(ByVal pDic As Scripting.Dictionary)
    
    Dim vCount As Integer

    Set oForm = frSoFifa
    
    If (pDic.Count = 0) Then
        Let oForm.txClubHidden.ListIndex = oForm.txClub.ListIndex
        GoTo Finish
    End If
    
    With oForm.txClubHidden
        Call .Clear
        For vCount = 0 To pDic.Count
            If (vCount = pDic.Count) Then Exit For
            Call .AddItem(pDic.Keys(vCount), vCount)
        Next vCount
        Let .ListIndex = oForm.txClub.ListIndex
    End With

Finish:
    Set oForm = Nothing
    Set pDic = Nothing

End Sub


Private Sub LoadPlayers(ByVal pDic As Scripting.Dictionary)

    Dim vCount As Integer

    Set oForm = frSoFifa

    With oForm.txPlayer
        Call .Clear
        For vCount = 0 To pDic.Count
            If (vCount = pDic.Count) Then Exit For
            Call .AddItem(pDic.Items(vCount), vCount)
        Next vCount
        Let .ListIndex = 0
    End With

    Set oForm = Nothing
    Set pDic = Nothing

End Sub
Private Sub LoadPlayersHidden(ByVal pDic As Scripting.Dictionary)

    Dim vCount As Integer

    Set oForm = frSoFifa

    With oForm.txPlayerHidden
        Call .Clear
        
        For vCount = 0 To pDic.Count
            If (vCount = pDic.Count) Then Exit For
            Call .AddItem(pDic.Keys(vCount), vCount)
        Next vCount
        Let .ListIndex = oForm.txPlayer.ListIndex
    End With

    Set oForm = Nothing
    Set pDic = Nothing

End Sub

