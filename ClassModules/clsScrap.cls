VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScrap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'https://www.pes6.es/stats/fifa-to-pes6
'#showManualStats
'#offlineStats
'#convertOffline
'#statsToCopy
'https://sofifa.com/player/157492/?r=080002&set=true

Private oSk As clsSoFifaSkill
Private oBs As clsSoFifaBasic
Private oClsGF As clsGlobalFunc
Private oDoc As MSHTML.HTMLDocument
Private oXml As MSXML2.XMLHTTP60


Private Function GetPlayerPage(ByVal pPageUrl As String) As MSHTML.HTMLDocument

Variables:
    Set oClsGF = New clsGlobalFunc
    Set oXml = New MSXML2.XMLHTTP60
    Set oDoc = New MSHTML.HTMLDocument

GetPage:
    On Error GoTo ErrHandler
    
    With oXml
        Call .Open("GET", pPageUrl, False)
        Call .send
        While (.readyState <> 4): DoEvents: Wend
        If (.Status <> 200) Then GoTo StatusErr
    End With
    
GetDocument:
    Let oDoc.body.innerHTML = oXml.responseText
    Let pPageUrl = Trim(Left(oDoc.body.innerHTML, 16))
    
    Do While (pPageUrl = "")
        Let oDoc.body.innerHTML = oXml.responseText
        Let pPageUrl = Trim(Left(oDoc.body.innerHTML, 16))
    Loop
    
    Set GetPlayerPage = oDoc
    
    On Error GoTo 0
    
    GoTo Finish
    
StatusErr:
    Call MsgBox(oXml.statusText, vbExclamation, oXml.Status)
    
Finish:
    On Error Resume Next
    Exit Function
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("Scrap", "GetPlayerPage")
    GoTo Finish

End Function


Public Function GetBasicInfo(ByVal pPageUrl As String) As Scripting.Dictionary

Variables:
    Set oBs = New clsSoFifaBasic
    Set oClsGF = New clsGlobalFunc
    Set oDoc = New MSHTML.HTMLDocument

GetDocument:
    Set oDoc = GetPlayerPage(pPageUrl)

ScrapPage:
    On Error GoTo ErrHandler
    
    With oDoc
        Let oBs.Name = .querySelector(".info > h1").innerText
        Let oBs.Nation = .querySelector(".info > .meta.ellipsis > a").Title
        Let oBs.Basics = .querySelector(".meta.ellipsis").childNodes.Item(3).Data
        Let oBs.Overall = .querySelectorAll("div.col.col-4 > ul span").Item(1).innerText
        Let oBs.PrefPos = .querySelectorAll("div.col.col-4 > ul span").Item(0).innerText
        Let oBs.Positions = .querySelector(".meta.ellipsis > span").innerText
        Let oBs.Foot = .querySelectorAll(".block-quarter ul li").Item(0).lastChild.Data
        Let oBs.WeakFoot = .querySelectorAll(".block-quarter ul li").Item(1).firstChild.Data
        Let oBs.SkillMoves = .querySelectorAll(".block-quarter ul li").Item(2).firstChild.Data
        Let oBs.InterRep = .querySelectorAll(".block-quarter ul li").Item(3).firstChild.Data
    End With

ReturningValue:
    Set GetBasicInfo = oBs.GetDictionary
        
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oBs = Nothing
    Set oClsGF = Nothing
    Set oDoc = Nothing
    Exit Function
        
ErrHandler:
    Call oClsGF.DisplayErrMessage("Scraping", "GetBasicInfo")
    GoTo Finish

End Function

Public Function GetSkills2007(ByVal pPageUrl As String) As Scripting.Dictionary
      
Variables:
    Dim vDoc As MSHTML.HTMLDocument
    Set oSk = New clsSoFifaSkill
    Set oClsGF = New clsGlobalFunc
    Set oDoc = New MSHTML.HTMLDocument
    Set vDoc = New MSHTML.HTMLDocument
    
GetDocument:
    Set oDoc = GetPlayerPage(pPageUrl)

ScrapPage:
    On Error GoTo ErrHandler
    
    Let vDoc.body.innerHTML = oDoc.querySelectorAll(".col-12").Item(1).innerHTML
    
    With vDoc.querySelectorAll(".bp3-tag")
        Let oSk.Crossing = .Item(0).innerText
        Let oSk.Finishing = .Item(1).innerText
        Let oSk.HeadingAcc = .Item(2).innerText
        Let oSk.ShortPass = .Item(3).innerText
        Let oSk.Dribbling = .Item(4).innerText
        Let oSk.FkAcc = .Item(5).innerText
        Let oSk.LongPass = .Item(6).innerText
        Let oSk.BallControl = .Item(7).innerText
        Let oSk.Acceleration = .Item(8).innerText
        Let oSk.SprintSpeed = .Item(9).innerText
        Let oSk.Reactions = .Item(10).innerText
        Let oSk.ShotPower = .Item(11).innerText
        Let oSk.Stamina = .Item(12).innerText
        Let oSk.Strength = .Item(13).innerText
        Let oSk.LongShots = .Item(14).innerText
        Let oSk.Aggression = .Item(15).innerText
        Let oSk.Penalties = .Item(16).innerText
        Let oSk.Marking = .Item(17).innerText
        Let oSk.StandingTackle = .Item(18).innerText
        Let oSk.GkDiving = .Item(19).innerText
        Let oSk.GkHandling = .Item(20).innerText
        Let oSk.GkKicking = .Item(21).innerText
        Let oSk.GkPositioning = .Item(22).innerText
        Let oSk.GkReflexes = .Item(23).innerText
        Let oSk.Interceptions = vbNullString
        Let oSk.Positioning = vbNullString
    End With
    
    If (vDoc.querySelectorAll("div.card").Item(7) Is Nothing) Then
        Let oSk.Traits = vbNullString
    Else
        Let oSk.Traits = vDoc.querySelectorAll("div.card").Item(7).innerText
    End If
    
ReturningValue:
    Set GetSkills2007 = oSk.GetDictionary
        
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oSk = Nothing
    Set oClsGF = Nothing
    Set oDoc = Nothing
    Exit Function
        
ErrHandler:
    Call oClsGF.DisplayErrMessage("Scraping", "GetSkills")
    GoTo Finish

End Function


Public Function GetSkills2008(ByVal pPageUrl As String) As Scripting.Dictionary
      
Variables:
    Dim vDoc As MSHTML.HTMLDocument
    Set oSk = New clsSoFifaSkill
    Set oClsGF = New clsGlobalFunc
    Set oDoc = New MSHTML.HTMLDocument
    Set vDoc = New MSHTML.HTMLDocument
    
GetDocument:
    Set oDoc = GetPlayerPage(pPageUrl)

ScrapPage:
    On Error GoTo ErrHandler
    
    Let vDoc.body.innerHTML = oDoc.querySelectorAll(".col-12").Item(1).innerHTML
    
    With vDoc.querySelectorAll(".bp3-tag")
        Let oSk.Crossing = .Item(0).innerText
        Let oSk.Finishing = .Item(1).innerText
        Let oSk.HeadingAcc = .Item(2).innerText
        Let oSk.ShortPass = .Item(3).innerText
        Let oSk.Dribbling = .Item(4).innerText
        Let oSk.FkAcc = .Item(5).innerText
        Let oSk.LongPass = .Item(6).innerText
        Let oSk.BallControl = .Item(7).innerText
        Let oSk.Acceleration = .Item(8).innerText
        Let oSk.SprintSpeed = .Item(9).innerText
        Let oSk.Reactions = .Item(10).innerText
        Let oSk.ShotPower = .Item(11).innerText
        Let oSk.Stamina = .Item(12).innerText
        Let oSk.Strength = .Item(13).innerText
        Let oSk.LongShots = .Item(14).innerText
        Let oSk.Aggression = .Item(15).innerText
        Let oSk.Interceptions = .Item(16).innerText
        Let oSk.Positioning = .Item(17).innerText
        Let oSk.Penalties = .Item(18).innerText
        Let oSk.Marking = .Item(19).innerText
        Let oSk.StandingTackle = .Item(20).innerText
        Let oSk.GkDiving = .Item(21).innerText
        Let oSk.GkHandling = .Item(22).innerText
        Let oSk.GkKicking = .Item(23).innerText
        Let oSk.GkPositioning = .Item(24).innerText
        Let oSk.GkReflexes = .Item(25).innerText
    End With
    
    If (vDoc.querySelectorAll("div.card").Item(7) Is Nothing) Then
        Let oSk.Traits = vbNullString
    Else
        Let oSk.Traits = vDoc.querySelectorAll("div.card").Item(7).innerText
    End If

ReturningValue:
    Set GetSkills2008 = oSk.GetDictionary
        
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oSk = Nothing
    Set oClsGF = Nothing
    Set oDoc = Nothing
    Exit Function
        
ErrHandler:
    Call oClsGF.DisplayErrMessage("Scraping", "GetSkills")
    GoTo Finish

End Function
