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


'Selenium VBA
'https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0

'ChromeDriver
'https://chromedriver.storage.googleapis.com/index.html?path=100.0.4896.60/


Public Sub SearchPlayerStats()
      
    Dim sb As MSComctlLib.StatusBar
    Dim frm As frSoFifa
    Dim cgf As New clsGlobalFunc
    Dim cos As New clsOldStats
    Dim dic As Scripting.Dictionary
    Dim json As String
    
    On Error GoTo ErrHandler
    
    Set frm = frSoFifa
    Set sb = frm.btStatusBar
    
    If (frm.txPlayerHidden.ListIndex = -1) Then
        Let sb.Panels(1).Text = "Player must be selected"
        GoTo Finish
    End If
    
    Let frm.txStatsToCopy.Text = vbNullString
    
    Select Case (frm.txGame.Text)
        Case "FIFA 07", "FIFA 08"
            Let sb.Panels(1).Text = "Parsing JSON"
            Let json = cos.GetPlayerOldStats()
            Let sb.Panels(1).Text = "Getting Pes6 offline stats (JSON)"
            Let frm.txStatsToCopy.Text = Pes6StatsSelenium(json)
        Case Else
            Let sb.Panels(1).Text = "Getting Pes6 stats (PLAYER ID)"
            Let frm.txStatsToCopy.Text = Pes6StatsSelenium(frm.txPlayerHidden.Text)
    End Select
    
    Let sb.Panels(1).Text = "Ready"
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set sb = Nothing
    Set frm = Nothing
    Exit Sub
    
ErrHandler:
    Let sb.Panels(1).Text = "Search incomplete!"
    Call cgf.DisplayErrMessage("Scrap", "SearchPlayerStats")
    GoTo Finish

End Sub


Public Function Pes6StatsSelenium(ByVal playerID As String) As String
    
    Dim ie As Selenium.ChromeDriver
    Dim frm As frSoFifa
    Dim sb As MSComctlLib.StatusBar
    Dim cgf As New clsGlobalFunc
    
    On Error GoTo ErrHandler
   
    Set ie = New Selenium.ChromeDriver
    Set frm = frSoFifa
    Set sb = frm.btStatusBar
        
    With ie
        Call .AddArgument("--headless")
        Let sb.Panels(1).Text = "Wait page load pes6.es"
        Call .get("https://www.pes6.es/stats/fifa-to-pes6.php")
        Call .Wait(1800)
        Let sb.Panels(1).Text = "Set playerID"
        
        Select Case (IsNumeric(playerID))
            Case True
                Call .FindElementByCss("#freeSoFifa").Clear
                Call .FindElementByCss("#freeSoFifa").SendKeys(playerID)
                Call .FindElementByCss("#convertFreeSoFifa").Click
            Case Else
                Call .FindElementByCss("#showManualStats").Click
                Call .FindElementByCss("#offlineStats").Clear
                Call .FindElementByCss("#offlineStats").SendKeys(playerID)
                Call .FindElementByCss("#convertOffline").Click
        End Select
        
        Let sb.Panels(1).Text = "Converting..."
        Call .Wait(2000)
    End With
    
    Let sb.Panels(1).Text = "Finished converting"
    Let Pes6StatsSelenium = ie.FindElementByCss("#statsToCopy").Attribute("innerText")
    
    Call ie.Quit

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set sb = Nothing
    Set frm = Nothing
    Set ie = Nothing
    Exit Function
    
ErrHandler:
    Let sb.Panels(1).Text = "Cannot get player stats"
    Call cgf.DisplayErrMessage("Scrap", "Pes6StatsSelenium")
    GoTo Finish
    
End Function
