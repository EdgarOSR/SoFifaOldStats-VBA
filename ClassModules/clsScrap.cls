VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOldStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const baseUrl As String = "https://sofifa.com/player/"


Public Function GetPlayerOldStats() As String
    
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    Dim dicSkill As Scripting.Dictionary
    Dim dicBasic As Scripting.Dictionary
    Dim json As String
    Dim url As String
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa
       
    Let url = baseUrl & frm.txPlayerHidden.Text & frm.txVersionHidden.Text
    
    Set dicBasic = GetBasicInfo(url)
    
    Select Case (frm.txGame.Text)
        Case "FIFA 07"
            Set dicSkill = GetSkills2007(url)
        Case "FIFA 08"
            Set dicSkill = GetSkills2008(url)
    End Select
    
    Let GetPlayerOldStats = GetJson(dicBasic, dicSkill)
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set frm = Nothing
    Set cgf = Nothing
    Set dicSkill = Nothing
    Set dicBasic = Nothing
    Exit Function
    
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "GetPlayerOldStats")
    GoTo Finish

End Function


Private Function GetPlayerPage(ByVal url As String) As MSHTML.HTMLDocument

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

    Set GetPlayerPage = doc

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set doc = Nothing
    Set xml = Nothing
    Set cgf = Nothing
    Exit Function

ErrHandler:
        Call cgf.DisplayErrMessage("OldStats", "GetPlayerPage")
    GoTo Finish

End Function


Private Function GetBasicInfo(ByVal pPageUrl As String) As Scripting.Dictionary

    Dim cbs As clsOldStatsBasic
    Dim cgf As clsGlobalFunc
    Dim doc As MSHTML.HTMLDocument
    
    Set cbs = New clsOldStatsBasic
    Set cgf = New clsGlobalFunc
    Set doc = New MSHTML.HTMLDocument

    Set doc = GetPlayerPage(pPageUrl)

    On Error GoTo ErrHandler
    
    With doc
        Let cbs.Name = .querySelector(".info > h1").innerText
        Let cbs.Nation = .querySelector(".info > .meta.ellipsis > a").Title
        Let cbs.Basics = .querySelector(".meta.ellipsis").childNodes.Item(3).Data
        Let cbs.Overall = .querySelectorAll("div.col.col-4 > ul span").Item(1).innerText
        Let cbs.PrefPos = .querySelectorAll("div.col.col-4 > ul span").Item(0).innerText
        Let cbs.Positions = .querySelector(".meta.ellipsis > span").innerText
        Let cbs.Foot = .querySelectorAll(".block-quarter ul li").Item(0).lastChild.Data
        Let cbs.WeakFoot = .querySelectorAll(".block-quarter ul li").Item(1).firstChild.Data
        Let cbs.SkillMoves = .querySelectorAll(".block-quarter ul li").Item(2).firstChild.Data
        Let cbs.InterRep = .querySelectorAll(".block-quarter ul li").Item(3).firstChild.Data
    End With

    Set GetBasicInfo = cbs.GetDictionary
        
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set cbs = Nothing
    Set cgf = Nothing
    Set doc = Nothing
    Exit Function
        
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "GetBasicInfo")
    GoTo Finish

End Function


Private Function GetSkills2007(ByVal pPageUrl As String) As Scripting.Dictionary
      
    Dim csk As clsOldStatsSkill
    Dim cgf As clsGlobalFunc
    Dim doc As MSHTML.HTMLDocument
    
    Set csk = New clsOldStatsSkill
    Set cgf = New clsGlobalFunc
    Set doc = New MSHTML.HTMLDocument
    
    Set doc = GetPlayerPage(pPageUrl)

    On Error GoTo ErrHandler
    
    Let doc.body.innerHTML = doc.querySelectorAll(".col-12").Item(1).innerHTML
    
    With doc.querySelectorAll(".bp3-tag")
        Let csk.Crossing = .Item(0).innerText
        Let csk.Finishing = .Item(1).innerText
        Let csk.HeadingAcc = .Item(2).innerText
        Let csk.ShortPass = .Item(3).innerText
        Let csk.Dribbling = .Item(4).innerText
        Let csk.FkAcc = .Item(5).innerText
        Let csk.LongPass = .Item(6).innerText
        Let csk.BallControl = .Item(7).innerText
        Let csk.Acceleration = .Item(8).innerText
        Let csk.SprintSpeed = .Item(9).innerText
        Let csk.Reactions = .Item(10).innerText
        Let csk.ShotPower = .Item(11).innerText
        Let csk.Stamina = .Item(12).innerText
        Let csk.Strength = .Item(13).innerText
        Let csk.LongShots = .Item(14).innerText
        Let csk.Aggression = .Item(15).innerText
        Let csk.Penalties = .Item(16).innerText
        Let csk.Marking = .Item(17).innerText
        Let csk.StandingTackle = .Item(18).innerText
        Let csk.GkDiving = .Item(19).innerText
        Let csk.GkHandling = .Item(20).innerText
        Let csk.GkKicking = .Item(21).innerText
        Let csk.GkPositioning = .Item(22).innerText
        Let csk.GkReflexes = .Item(23).innerText
        Let csk.Interceptions = vbNullString
        Let csk.Positioning = vbNullString
    End With
    
    If (doc.querySelectorAll("div.card").Item(7) Is Nothing) Then
        Let csk.Traits = vbNullString
    Else
        Let csk.Traits = doc.querySelectorAll("div.card").Item(7).innerText
    End If

    Set GetSkills2007 = csk.GetDictionary
        
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set csk = Nothing
    Set cgf = Nothing
    Set doc = Nothing
    Exit Function
        
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "GetSkills2007")
    GoTo Finish

End Function


Public Function GetSkills2008(ByVal pPageUrl As String) As Scripting.Dictionary
      
    Dim csk As clsOldStatsSkill
    Dim cgf As clsGlobalFunc
    Dim doc As MSHTML.HTMLDocument
    
    Set csk = New clsOldStatsSkill
    Set cgf = New clsGlobalFunc
    Set doc = New MSHTML.HTMLDocument
    
    Set doc = GetPlayerPage(pPageUrl)

    On Error GoTo ErrHandler
    
    Let doc.body.innerHTML = doc.querySelectorAll(".col-12").Item(1).innerHTML
    
    With doc.querySelectorAll(".bp3-tag")
        Let csk.Crossing = .Item(0).innerText
        Let csk.Finishing = .Item(1).innerText
        Let csk.HeadingAcc = .Item(2).innerText
        Let csk.ShortPass = .Item(3).innerText
        Let csk.Dribbling = .Item(4).innerText
        Let csk.FkAcc = .Item(5).innerText
        Let csk.LongPass = .Item(6).innerText
        Let csk.BallControl = .Item(7).innerText
        Let csk.Acceleration = .Item(8).innerText
        Let csk.SprintSpeed = .Item(9).innerText
        Let csk.Reactions = .Item(10).innerText
        Let csk.ShotPower = .Item(11).innerText
        Let csk.Stamina = .Item(12).innerText
        Let csk.Strength = .Item(13).innerText
        Let csk.LongShots = .Item(14).innerText
        Let csk.Aggression = .Item(15).innerText
        Let csk.Interceptions = .Item(16).innerText
        Let csk.Positioning = .Item(17).innerText
        Let csk.Penalties = .Item(18).innerText
        Let csk.Marking = .Item(19).innerText
        Let csk.StandingTackle = .Item(20).innerText
        Let csk.GkDiving = .Item(21).innerText
        Let csk.GkHandling = .Item(22).innerText
        Let csk.GkKicking = .Item(23).innerText
        Let csk.GkPositioning = .Item(24).innerText
        Let csk.GkReflexes = .Item(25).innerText
    End With
    
    If (doc.querySelectorAll("div.card").Item(7) Is Nothing) Then
        Let csk.Traits = vbNullString
    Else
        Let csk.Traits = doc.querySelectorAll("div.card").Item(7).innerText
    End If

    Set GetSkills2008 = csk.GetDictionary
        
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set csk = Nothing
    Set cgf = Nothing
    Set doc = Nothing
    Exit Function
        
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "GetSkills2008")
    GoTo Finish

End Function


Private Function GetJson(ByVal dicBasic As Scripting.Dictionary, ByVal dicSkill As Scripting.Dictionary) As String

    Dim cgf As clsGlobalFunc
    Dim json As String
    
    Set cgf = New clsGlobalFunc
    Let json = vbNullString

ParsingJSON:
    Let json = json & JsonParseBasics(dicBasic)
    Let json = json & JsonParseSkills(dicSkill)
    
ReturningValue:
    Let GetJson = json
    
Finish:
    On Error Resume Next
    Set cgf = Nothing
    Exit Function
    
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "GetJson")
    GoTo Finish

End Function


Private Function JsonParseBasics(ByVal pDic As Scripting.Dictionary) As String

    Dim cgf As clsGlobalFunc
    Dim key As Variant
    Dim json As String
    
    Set cgf = New clsGlobalFunc
    
    Let json = vbNullString
    Let json = Chr(123)
    
    For Each key In pDic.Keys
        If (IsEmpty(key)) Then
            Exit For
        ElseIf (key = "positions") Then
            Let json = json & Chr(34) & key & Chr(34) & Chr(58) & pDic(key)
        ElseIf (IsNumeric(pDic(key))) Then
            Let json = json & Chr(34) & key & Chr(34) & Chr(58) & pDic(key)
        Else
            Let json = json & Chr(34) & key & Chr(34) & Chr(58) & Chr(34) & pDic(key) & Chr(34)
        End If
        Let json = json & Chr(44)
    Next key
    
    Let JsonParseBasics = json

Finish:
    On Error Resume Next
    Set cgf = Nothing
    Set key = Nothing
    Exit Function
    
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "JsonParseBasics")
    GoTo Finish

End Function


Private Function JsonParseSkills(ByVal pDic As Scripting.Dictionary) As String

    Dim cgf As clsGlobalFunc
    Dim key As Variant
    Dim json As String
    
    Set cgf = New clsGlobalFunc
    
    Let json = vbNullString
    Let json = Chr(34) & "skills" & Chr(34) & Chr(58) & Chr(123)
    
    For Each key In pDic.Keys
        If (IsEmpty(key)) Then
            Exit For
        ElseIf (key = "player_traits") Then
            Let json = Left(json, Len(json) - 1)
            Let json = json & Chr(125) & Chr(44)
            Let json = json & Chr(34) & key & Chr(34) & Chr(58) & pDic(key) & Chr(125)
            Exit For
        ElseIf (IsNumeric(pDic(key))) Then
            Let json = json & Chr(34) & key & Chr(34) & Chr(58) & pDic(key)
        Else
            Let json = json & Chr(34) & key & Chr(34) & Chr(58) & Chr(34) & pDic(key) & Chr(34)
        End If
        Let json = json & Chr(44)
    Next key
    
    Let JsonParseSkills = json

Finish:
    On Error Resume Next
    Set cgf = Nothing
    Set key = Nothing
    Exit Function
    
ErrHandler:
    Call cgf.DisplayErrMessage("OldStats", "JsonParseSkills")
    GoTo Finish

End Function
