VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoFifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const mBaseUrl As String = "https://sofifa.com/player/@id@/?r=@version@&set=true"
Private oClsGF As clsGlobalFunc
Private oClsSC As clsScrap
Private oClsPJ As clsParseJson
Private oSb As MSComctlLib.StatusBar
Private oForm As frSoFifa

Public Sub SearchPlayerStats()
   
    Dim oDoc As MSHTML.HTMLDocument
    Dim oDicBasics As Scripting.Dictionary
    Dim oDicSkills As Scripting.Dictionary
    Dim vUrl As String

    Set oForm = frSoFifa
    If (oForm.txPlayerHidden.ListIndex = -1) Then
        Call MsgBox("Player must be selected", vbExclamation)
        GoTo Finish
    End If

    Set oClsGF = New clsGlobalFunc
    Set oClsSC = New clsScrap
    Set oClsPJ = New clsParseJson
    Set oDoc = New MSHTML.HTMLDocument
    Set oDicBasics = New Scripting.Dictionary
    Set oDicSkills = New Scripting.Dictionary
    Set oSb = oForm.btStatusBar
    Let vUrl = mBaseUrl
    Let vUrl = Replace(vUrl, "@id@", oForm.txPlayerHidden.Text)
    Let vUrl = Replace(vUrl, "@version@", oForm.txVersion.Text)
    
GetValues:
    On Error GoTo ErrHandler
    
    Let oSb.Panels(1).Text = "SCRAPING BASIC INFO"
    Set oDicBasics = oClsSC.GetBasicInfo(vUrl)
    
    Let oSb.Panels(1).Text = "SCRAPING SKILLS"
    Select Case (oForm.txVersion.Text)
        Case "070001", "070002"
            Set oDicSkills = oClsSC.GetSkills2007(vUrl)
        Case Else
            Set oDicSkills = oClsSC.GetSkills2008(vUrl)
    End Select

    Let oSb.Panels(1).Text = "PARSING JSON"
    Let vUrl = oClsPJ.GetJSON(oDicBasics, oDicSkills)

    Let oSb.Panels(1).Text = "GETTING PES6 STATS"
    Let oForm.txStatsToCopy.Text = Pes6StatsSelenium(vUrl)
    
    Let oSb.Panels(1).Text = "READY"
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oSb = Nothing
    Set oDoc = Nothing
    Set oDicBasics = Nothing
    Set oDicSkills = Nothing
    Set oForm = Nothing
    Let vUrl = vbNullString
    Exit Sub
    
ErrHandler:
    Let oSb.Panels(1).Text = "EXIT WITH ERROR"
    Call oClsGF.DisplayErrMessage("Scrap", "GetPlayerPage")
    GoTo Finish

End Sub


Public Function Pes6StatsSelenium(ByVal pJSON As String) As String

'Selenium VBA
'https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0

'ChromeDriver
'https://chromedriver.storage.googleapis.com/index.html?path=100.0.4896.60/

    Dim vResponse As String
    Dim oIe As Selenium.ChromeDriver
    Set oIe = New Selenium.ChromeDriver
    Let vResponse = vbNullString

    With oIe
        Call .AddArgument("--headless")
        Call .get("https://www.pes6.es/stats/fifa-to-pes6")
        Call .Wait(1800)
        Call .FindElementByCss("#showManualStats").Click
        Call .FindElementByCss("#offlineStats").Clear
        Call .FindElementByCss("#offlineStats").SendKeys(pJSON)
        Call .FindElementByCss("#convertOffline").Click
        Call .Wait(1800)
    End With
    
    Let vResponse = oIe.FindElementByCss("#statsToCopy").Attribute("innerText")
    
    Call oIe.Quit
    
    Let Pes6StatsSelenium = vResponse
    
    Set oIe = Nothing

End Function


Private Sub DebugDict(ByVal pDic As Dictionary)
    Dim k As Variant
    For Each k In pDic.Keys
        Debug.Print k & " --> " & pDic.Item(k)
    Next k
    Set k = Nothing
    Set pDic = Nothing
End Sub
