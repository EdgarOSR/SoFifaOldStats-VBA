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


Private Const baseUrl As String = "https://sofifa.com/"


Public Sub SearchTeamsList()
    
    Dim dic As Scripting.Dictionary
    Dim doc As MSHTML.HTMLDocument
    Dim url As String
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa
    Set doc = New MSHTML.HTMLDocument
    
    Let url = baseUrl & "/teams/" & frm.txVersionHidden.Text
    Let url = url & frm.txCountryHidden.Text
    Let url = url & "&col=oa&sort=desc"

    Set doc = GetDocHtml(url)
    Set dic = GetTeamsList(doc)
    Call LoadTeams(dic)
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set doc = Nothing
    Set dic = Nothing
    Set frm = Nothing
    Set cgf = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "SearchTeamsList")
    GoTo Finish

End Sub


Public Sub SearchPlayersList()
    
    Dim dic As Scripting.Dictionary
    Dim doc As MSHTML.HTMLDocument
    Dim url As String
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa
    Set doc = New MSHTML.HTMLDocument
    
    Select Case (frm.txTeamHidden.Text)
        Case -1
            Let url = baseUrl & "players/" & frm.txVersionHidden.Text
            Let url = url & frm.txCountryHidden.Text
            Let url = url & "&oal=" & frm.txMinOver.Value & "&oah=" & frm.txMaxOver.Value
            If (frm.txType.Text <> "All") Then
                Let url = url & "&type=" & LCase(frm.txType.Text)
            End If
        Case Else
            Let url = baseUrl & "team/@team@/" & frm.txVersionHidden
            Let url = Replace(url, "@team@", frm.txTeamHidden.Text)
    End Select
    
    If (frm.txOrderClub = True) Then
        Let url = url & "&col=tm&sort=asc&offset=0"
    Else
        Let url = url & "&col=oa&sort=desc&offset=0"
    End If
    
    Set doc = GetDocHtml(url)
    
    Select Case (frm.txTeamHidden.Text)
        Case -1
            If (frm.txType.Text = "Updated") Then
                Set dic = GetPlayersListUpdatedByNation(doc)
            Else
                Set dic = GetPlayersListByNation(doc)
            End If
        Case Else
            Set dic = GetPlayersListByTeam(doc)
    End Select
    
    Call LoadPlayers(dic)
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set doc = Nothing
    Set dic = Nothing
    Set frm = Nothing
    Set cgf = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "SearchPlayersList")
    GoTo Finish

End Sub


Private Function GetDocHtml(ByVal url As String) As MSHTML.HTMLDocument

    Dim xml As MSXML2.XMLHTTP60
    Dim doc As MSHTML.HTMLDocument
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set xml = New MSXML2.XMLHTTP60
    Set doc = New MSHTML.HTMLDocument

    With xml
        Call .Open("GET", url, False)
        Call .send
        Do While (.readyState <> 4): DoEvents: Loop
        If (.Status <> 200) Then GoTo ErrHandler
    End With

    Let doc.body.innerHTML = xml.responseText

    Set GetDocHtml = doc

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set doc = Nothing
    Set xml = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
        Call cgf.DisplayErrMessage("SoFifaLeagues", "GetDocHtml")
    GoTo Finish

End Function


Private Function GetTeamsList(ByVal doc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    Dim i As Integer
    Dim max As Integer
    Dim id As Double
    Dim aux As String
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
        
    Set dic = New Scripting.Dictionary

    Let max = doc.querySelectorAll(".card .col-name-wide a").length - 2
    
    Call dic.Add(-1, "None (Player By Nation)")
    
    For i = 0 To max Step 2
        With doc.querySelectorAll(".card .col-name-wide a")
            Let id = 0
            Let id = Split(.Item(i).href, "/")(2)
            Let aux = vbNullString
            Let aux = .Item(i).innerText
        End With
        Call dic.Add(id, aux)
    Next i

    Set GetTeamsList = dic

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set dic = Nothing
    Set doc = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "GetTeamsList")
    GoTo Finish

End Function


Private Function FixPlayerName(ByVal Name As String) As String

    Dim arr As Variant
    Let arr = Split(Name, "-")
    
    Let FixPlayerName = arr(LBound(arr)) & Chr(32) & arr(UBound(arr))
    Let FixPlayerName = StrConv(FixPlayerName, vbProperCase)

End Function


Private Function GetPlayersListByNation(ByVal doc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    Dim chd As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
    Dim aux As String
    Dim info As String
    Dim id As Double
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
        
    Set dic = New Scripting.Dictionary

GetPlayers:
    Set chd = doc.querySelectorAll("tbody tr")
    Do While i < chd.length
        Let aux = vbNullString
        Let info = vbNullString
        With chd.Item(i)
            Let aux = Trim(.children.Item(3).children.Item(0).innerText)
            Let info = info & aux & String(3 - Len(aux), Chr(32)) & Chr(9)
            Let aux = Trim(.children.Item(1).children.Item(2).innerText)
            Let info = info & aux & String(3 - Len(aux), Chr(32)) & Chr(9)
            Let aux = Trim(.children.Item(1).children.Item(0).innerText)
            Let info = info & aux & String(29 - Len(aux), Chr(32)) & Chr(9)
            Let aux = Trim(Split(.children.Item(5).innerText, vbCrLf)(0))
            Let info = info & aux
            Let id = Val(Split(.children.Item(1).children.Item(0).href, "/")(2))
            Call dic.Add(id, info)
        End With
        Let i = i + 1
    Loop
    
    Set chd = doc.querySelectorAll(".pagination .bp3-icon-chevron-right")
    If (chd.length > 0) Then
        Let aux = Replace(chd.Item(0).parentElement.href, "about:", "https://sofifa.com")
        Set doc = GetDocHtml(aux)
        Let i = 0
        GoTo GetPlayers
    End If

    Set GetPlayersListByNation = dic

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set dic = Nothing
    Set doc = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "GetPlayersList")
    GoTo Finish

End Function


Private Function GetPlayersListUpdatedByNation(ByVal doc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    Dim chd As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
    Dim aux As String
    Dim info As String
    Dim id As Double
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
        
    Set dic = New Scripting.Dictionary

GetPlayers:
    Set chd = doc.querySelectorAll("tbody tr")
    Do While i < chd.length
        Let aux = vbNullString
        Let info = vbNullString
        On Error Resume Next
            With chd.Item(i)
                Let aux = Trim(.children.Item(0).innerText)
                Let aux = Trim(Split(Split(aux, Chr(13))(2))(0))
                If (IsNumeric(aux)) Then
                    Let aux = Trim(.children.Item(3).children.Item(0).innerText)
                    Let info = info & aux & String(3 - Len(aux), Chr(32)) & Chr(9)
                    Let aux = Trim(.children.Item(1).children.Item(2).innerText)
                    Let info = info & aux & String(3 - Len(aux), Chr(32)) & Chr(9)
                    Let aux = Trim(.children.Item(1).children.Item(0).innerText)
                    Let info = info & aux & String(29 - Len(aux), Chr(32)) & Chr(9)
                    Let aux = Trim(Split(.children.Item(5).innerText, vbCrLf)(0))
                    Let info = info & aux
                    Let id = Val(Split(.children.Item(1).children.Item(0).href, "/")(2))
                    Call dic.Add(id, info)
                End If
            End With
        On Error GoTo 0
        Let i = i + 1
    Loop
    
    Set chd = doc.querySelectorAll(".pagination .bp3-icon-chevron-right")
    If (chd.length > 0) Then
        Let aux = Replace(chd.Item(0).parentElement.href, "about:", "https://sofifa.com")
        Set doc = GetDocHtml(aux)
        Let i = 0
        GoTo GetPlayers
    End If

    Set GetPlayersListUpdatedByNation = dic

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set dic = Nothing
    Set doc = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "GetPlayersList")
    GoTo Finish

End Function


Private Function GetPlayersListByTeam(ByVal doc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    Dim chd As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
    Dim aux As String
    Dim info As String
    Dim id As Double
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
        
    Set dic = New Scripting.Dictionary
    
    Set chd = doc.querySelectorAll("tbody tr")

    Do While i < chd.length
        Let aux = vbNullString
        Let info = vbNullString
        With chd.Item(i)
            Let aux = Trim(.children.Item(3).children.Item(0).innerText)
            Let info = info & aux & String(3 - Len(aux), Chr(32)) & Chr(9)
            Let aux = Trim(.children.Item(1).children.Item(2).innerText)
            Let info = info & aux & String(3 - Len(aux), Chr(32)) & Chr(9)
            Let aux = Trim(.children.Item(1).children.Item(0).innerText)
'            Let aux = Split(.children.Item(1).children.Item(0).href, "/")(3)
'            Let aux = FixPlayerName(aux)
            Let info = info & aux & String(29 - Len(aux), Chr(32)) & Chr(9)
            Let aux = Trim(.children.Item(5).children.Item(0).childNodes.Item(1).Data)
            Let info = info & aux
            Let id = Val(Split(.children.Item(1).children.Item(0).href, "/")(2))
        End With
        Call dic.Add(id, info)
        Let i = i + 1
    Loop

    Set GetPlayersListByTeam = dic

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set dic = Nothing
    Set doc = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "GetPlayersList")
    GoTo Finish

End Function


Private Sub LoadTeams(ByVal dic As Scripting.Dictionary)
    
    Dim i As Integer
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa

    Call frm.txTeam.Clear
    Call frm.txTeamHidden.Clear
    
    If (dic.Count = 0) Then
        Let frm.txTeam.ListIndex = -1
        Let frm.txTeamHidden.ListIndex = -1
        GoTo Finish
    End If
    
    Do While i < dic.Count
        Call frm.txTeam.AddItem(dic.Items(i), i)
        Call frm.txTeamHidden.AddItem(dic.Keys(i), i)
        Let i = i + 1
    Loop
    
    Let frm.txTeam.ListIndex = 0
    Let frm.txTeamHidden.ListIndex = frm.txTeam.ListIndex
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set frm = Nothing
    Set dic = Nothing
    Set cgf = Nothing
    Exit Sub
    
ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "LoadTeams")
    GoTo Finish

End Sub




Private Sub LoadPlayers(ByVal dic As Scripting.Dictionary)

    Dim i As Integer
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa

    Call frm.listPlayer.Clear
    Call frm.txPlayerHidden.Clear

    If (dic.Count = 0) Then
        Let frm.listPlayer.ListIndex = -1
        Let frm.txPlayerHidden.ListIndex = -1
        GoTo Finish
    End If

    Do While i < dic.Count
        Call frm.listPlayer.AddItem(dic.Items(i), i)
        Call frm.txPlayerHidden.AddItem(dic.Keys(i), i)
        Let i = i + 1
    Loop
        
    Let frm.listPlayer.ListIndex = 0
    Let frm.txPlayerHidden.ListIndex = frm.listPlayer.ListIndex

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set frm = Nothing
    Set dic = Nothing
    Set cgf = Nothing
    Exit Sub
    
ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "LoadPlayers")
    GoTo Finish

End Sub





Private Function GetGamesList(ByVal doc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    Dim chd As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
    Dim aux As String
    Dim id As Double
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
        
    Set dic = New Scripting.Dictionary
    
    Set chd = doc.querySelectorAll(".dropdown")
    
    If (chd.length = 0) Then GoTo Finish
    
    Let i = 0
    
    Do While i < chd.Item(1).children.Item(1).children.length
        With chd.Item(1).children.Item(1).children.Item(i)
            Call dic.Add(.innerText, Replace(.href, "about:/", vbNullString))
        End With
        Let i = i + 1
    Loop
    
    Set GetGamesList = dic

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set dic = Nothing
    Set doc = Nothing
    Set chd = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "GetGamesList")
    GoTo Finish

End Function


Private Function GetVersionsList(ByVal doc As MSHTML.HTMLDocument) As Scripting.Dictionary

    Dim dic As Scripting.Dictionary
    Dim chd As MSHTML.IHTMLDOMChildrenCollection
    Dim i As Integer
    Dim aux As String
    Dim id As Double
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
        
    Set dic = New Scripting.Dictionary
    
    Set chd = doc.querySelectorAll(".dropdown")
    
    If (chd.length = 0) Then GoTo Finish
    
    Let i = 0
    
    Do While i < chd.Item(2).children.Item(1).children.length
        With chd.Item(2).children.Item(1).children.Item(i)
            Call dic.Add(.innerText, Replace(.href, "about:/", vbNullString))
        End With
        Let i = i + 1
    Loop

    Set GetVersionsList = dic

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set dic = Nothing
    Set doc = Nothing
    Set chd = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "GetVersionsList")
    GoTo Finish

End Function


Private Sub LoadGames(ByVal dic As Scripting.Dictionary)
    
    Dim i As Integer
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa

    Call frm.txGame.Clear
    Call frm.txGameHidden.Clear

    If (dic.Count = 0) Then
        Let frm.txGame.ListIndex = -1
        Let frm.txGameHidden.ListIndex = frm.txGameHidden.ListIndex
        GoTo Finish
    End If
    
    Do While i < dic.Count
        Call frm.txGame.AddItem(dic.Keys(i), i)
        Call frm.txGameHidden.AddItem(dic.Items(i), i)
        Let i = i + 1
    Loop
    
    Let frm.txGame.ListIndex = 0
    Let frm.txGameHidden.ListIndex = frm.txGame.ListIndex

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set frm = Nothing
    Set dic = Nothing
    Set cgf = Nothing
    Exit Sub
    
ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "LoadGames")
    GoTo Finish

End Sub


Private Sub LoadVersions(ByVal dic As Scripting.Dictionary)
    
    Dim i As Integer
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa

    If (dic.Count = 0) Then
        Let frm.txVersion.ListIndex = -1
        Let frm.txVersionHidden.ListIndex = frm.txVersion.ListIndex
        GoTo Finish
    End If
    
    Call frm.txVersion.Clear
    Call frm.txVersionHidden.Clear
    
    Do While i < dic.Count
        Call frm.txVersion.AddItem(dic.Keys(i), i)
        Call frm.txVersionHidden.AddItem(dic.Items(i), i)
        Let i = i + 1
    Loop
    
    Let frm.txVersion.ListIndex = 0
    Let frm.txVersionHidden.ListIndex = frm.txVersion.ListIndex

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set frm = Nothing
    Set dic = Nothing
    Set cgf = Nothing
    Exit Sub
    
ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "LoadVersions")
    GoTo Finish

End Sub


Public Sub SearchGameList()
    
    Dim dic As Scripting.Dictionary
    Dim doc As MSHTML.HTMLDocument
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa
    Set doc = New MSHTML.HTMLDocument

    Set doc = GetDocHtml(baseUrl)
    Set dic = GetGamesList(doc)
    Call LoadGames(dic)

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set doc = Nothing
    Set dic = Nothing
    Set frm = Nothing
    Set cgf = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "SearchTeamsList")
    GoTo Finish

End Sub


Public Sub SearchVersionList()
    
    Dim dic As Scripting.Dictionary
    Dim doc As MSHTML.HTMLDocument
    Dim url As String
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa
    Set doc = New MSHTML.HTMLDocument
    
    Let url = baseUrl & frm.txGameHidden.Text

    Set doc = GetDocHtml(url)
    Set dic = GetVersionsList(doc)
    Call LoadVersions(dic)

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set doc = Nothing
    Set dic = Nothing
    Set frm = Nothing
    Set cgf = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaLeagues", "SearchVersionList")
    GoTo Finish

End Sub
