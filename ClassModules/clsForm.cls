VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private frm As frSoFifa
Private cgf As clsGlobalFunc
Private oClsSF As clsSoFifa
Private oClsSFL As clsSoFifaLeagues


Public Function FormClose() As Integer

    Let FormClose = False

    Set cgf = New clsGlobalFunc
    Set frm = frSoFifa
    
    On Error GoTo ErrHandler
    
    Select Case (MsgBox("Do you really like to exit?", vbYesNo, "Exit"))
        Case vbNo
            Let FormClose = True
            GoTo Finish
        Case Else
    End Select

    Call frm.Hide
    Call Unload(frm)

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set cgf = Nothing
    Set frm = Nothing
    Exit Function
    
ErrHandler:
    Call cgf.DisplayErrMessage("clsForm", "FormClose")
    GoTo Finish
    
End Function


Public Sub BtnCopyClick()
    
    Dim oDto As VB.Clipboard
    
    Set cgf = New clsGlobalFunc
    Set oDto = VB.Clipboard
    Set frm = frSoFifa
    
    On Error GoTo ErrHandler
    
    With frm
        If (.txStatsToCopy.Text = vbNullString) Then GoTo Finish
        Call oDto.Clear
        Let .txStatsToCopy.Text = Replace(.txStatsToCopy.Text, "?", "*")
        Call oDto.SetText(.txStatsToCopy.Text)
        Call MsgBox("Paste it in PES6 Editor", vbInformation)
    End With
    
    On Error GoTo 0

Finish:
    On Error Resume Next
    Set cgf = Nothing
    Set frm = Nothing
    Set oDto = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("clsForm", "BtnCopyClick")
    GoTo Finish
    
    
End Sub


Public Sub BtnClearClick()
    
    Set cgf = New clsGlobalFunc
    Set frm = frSoFifa
    
    On Error GoTo ErrHandler

    With frm
        Call frm.listPlayer.Clear
        Call frm.txPlayerHidden.Clear
        Call frm.txTeam.Clear
        Call frm.txTeamHidden.Clear
        Let frm.txMinOver.SelStart = 70
        Let frm.txMaxOver.SelStart = 99
        Call frm.txVersion.SetFocus
        Let frm.txStatsToCopy.Text = vbNullString
        Let frm.txCountry.ListIndex = 5
        Let frm.txCountryHidden.ListIndex = 5
        
    End With
    
    Call SearchTeamsList
    Call SearchPlayersList
    
    On Error GoTo 0

Finish:
    On Error Resume Next
    Set frm = Nothing
    Set cgf = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("clsForm", "BtnClearClick")
    GoTo Finish
    
End Sub


Public Sub FormActivate()
    
    Call SearchGameList
    Call SearchVersionList
    Call LoadCountries
    Call LoadCountriesHidden
    Call SearchTeamsList
    Call SearchPlayersList

End Sub


Private Sub LoadCountries()

    Set cgf = New clsGlobalFunc
    Set frm = frSoFifa

    On Error GoTo ErrHandler

    With frm.txCountry
        Call .Clear
        Call .AddItem("Austria", 0)
        Call .AddItem("Belgium", 1)
        Call .AddItem("Croatia", 2)
        Call .AddItem("Czech Republic", 3)
        Call .AddItem("Denmark", 4)
        Call .AddItem("England", 5)
        Call .AddItem("France", 6)
        Call .AddItem("Germany", 7)
        Call .AddItem("Greece", 8)
        Call .AddItem("Hungary", 9)
        Call .AddItem("Italy", 10)
        Call .AddItem("Netherlands", 11)
        Call .AddItem("Norway", 12)
        Call .AddItem("Poland", 13)
        Call .AddItem("Portugal", 14)
        Call .AddItem("Romania", 15)
        Call .AddItem("Russia", 16)
        Call .AddItem("Scotland", 17)
        Call .AddItem("Spain", 18)
        Call .AddItem("Sweden", 19)
        Call .AddItem("Switzerland", 20)
        Call .AddItem("Turkey", 21)
        Call .AddItem("Ukraine", 22)
        Call .AddItem("Argentina", 23)
        Call .AddItem("Bolivia", 24)
        Call .AddItem("Brazil", 25)
        Call .AddItem("Chile", 26)
        Call .AddItem("Colombia", 27)
        Call .AddItem("Ecuador", 28)
        Call .AddItem("Paraguay", 29)
        Call .AddItem("Peru", 30)
        Call .AddItem("Uruguay", 31)
        Call .AddItem("Venezuela", 32)
        Call .AddItem("Canada", 33)
        Call .AddItem("Mexico", 34)
        Call .AddItem("United States", 35)
        Call .AddItem("China PR", 36)
        Call .AddItem("Japan", 37)
        Call .AddItem("Saudi Arabia", 38)
        Call .AddItem("Australia", 39)
        Call .AddItem("Albania", 40)
        Call .AddItem("Bosnia", 41)
        Call .AddItem("Finland", 42)
        Call .AddItem("North Macedonia", 43)
        Call .AddItem("Iceland", 44)
        Call .AddItem("Ireland", 45)
        Call .AddItem("Northern Ireland", 46)
        Call .AddItem("Slovakia", 47)
        Call .AddItem("Slovenia", 48)
        Call .AddItem("Wales", 49)
        Call .AddItem("Serbia", 50)
        Call .AddItem("Costa Rica", 51)
        Call .AddItem("Jamaica", 52)
        Call .AddItem("Panama", 53)
        Call .AddItem("Algeria", 54)
        Call .AddItem("Cameroon", 55)
        Call .AddItem("Ivory Coast", 56)
        Call .AddItem("Egypt", 57)
        Call .AddItem("Ghana", 58)
        Call .AddItem("Morocco", 59)
        Call .AddItem("Nigeria", 60)
        Call .AddItem("Senegal", 61)
        Call .AddItem("Tunisia", 62)
        Call .AddItem("Iran", 63)
        Call .AddItem("Korea Republic", 64)
        Call .AddItem("Qatar", 65)
        Call .AddItem("United Arab Emirates", 66)
        Call .AddItem("New Zealand", 67)
        Let .ListIndex = 5
    End With

    On Error GoTo 0

Finish:
    On Error Resume Next
    Set cgf = Nothing
    Set frm = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("clsForm", "LoadCountries")
    GoTo Finish

End Sub


Private Sub LoadCountriesHidden()

    Set cgf = New clsGlobalFunc
    Set frm = frSoFifa

    On Error GoTo ErrHandler

    With frm.txCountryHidden
        Call .Clear
        Call .AddItem("&na%5B%5D=4", 0)
        Call .AddItem("&na%5B%5D=7", 1)
        Call .AddItem("&na%5B%5D=10", 2)
        Call .AddItem("&na%5B%5D=12", 3)
        Call .AddItem("&na%5B%5D=13", 4)
        Call .AddItem("&na%5B%5D=14", 5)
        Call .AddItem("&na%5B%5D=18", 6)
        Call .AddItem("&na%5B%5D=21", 7)
        Call .AddItem("&na%5B%5D=22", 8)
        Call .AddItem("&na%5B%5D=23", 9)
        Call .AddItem("&na%5B%5D=27", 10)
        Call .AddItem("&na%5B%5D=34", 11)
        Call .AddItem("&na%5B%5D=36", 12)
        Call .AddItem("&na%5B%5D=37", 13)
        Call .AddItem("&na%5B%5D=38", 14)
        Call .AddItem("&na%5B%5D=39", 15)
        Call .AddItem("&na%5B%5D=40", 16)
        Call .AddItem("&na%5B%5D=42", 17)
        Call .AddItem("&na%5B%5D=45", 18)
        Call .AddItem("&na%5B%5D=46", 19)
        Call .AddItem("&na%5B%5D=47", 20)
        Call .AddItem("&na%5B%5D=48", 21)
        Call .AddItem("&na%5B%5D=49", 22)
        Call .AddItem("&na%5B%5D=52", 23)
        Call .AddItem("&na%5B%5D=53", 24)
        Call .AddItem("&na%5B%5D=54", 25)
        Call .AddItem("&na%5B%5D=55", 26)
        Call .AddItem("&na%5B%5D=56", 27)
        Call .AddItem("&na%5B%5D=57", 28)
        Call .AddItem("&na%5B%5D=58", 29)
        Call .AddItem("&na%5B%5D=59", 30)
        Call .AddItem("&na%5B%5D=60", 31)
        Call .AddItem("&na%5B%5D=61", 32)
        Call .AddItem("&na%5B%5D=70", 33)
        Call .AddItem("&na%5B%5D=83", 34)
        Call .AddItem("&na%5B%5D=95", 35)
        Call .AddItem("&na%5B%5D=155", 36)
        Call .AddItem("&na%5B%5D=163", 37)
        Call .AddItem("&na%5B%5D=183", 38)
        Call .AddItem("&na%5B%5D=195", 39)
        Call .AddItem("&na%5B%5D=1", 40)
        Call .AddItem("&na%5B%5D=8", 41)
        Call .AddItem("&na%5B%5D=17", 42)
        Call .AddItem("&na%5B%5D=19", 43)
        Call .AddItem("&na%5B%5D=24", 44)
        Call .AddItem("&na%5B%5D=25", 45)
        Call .AddItem("&na%5B%5D=35", 46)
        Call .AddItem("&na%5B%5D=43", 47)
        Call .AddItem("&na%5B%5D=44", 48)
        Call .AddItem("&na%5B%5D=50", 49)
        Call .AddItem("&na%5B%5D=51", 50)
        Call .AddItem("&na%5B%5D=72", 51)
        Call .AddItem("&na%5B%5D=82", 52)
        Call .AddItem("&na%5B%5D=87", 53)
        Call .AddItem("&na%5B%5D=97", 54)
        Call .AddItem("&na%5B%5D=103", 55)
        Call .AddItem("&na%5B%5D=108", 56)
        Call .AddItem("&na%5B%5D=111", 57)
        Call .AddItem("&na%5B%5D=117", 58)
        Call .AddItem("&na%5B%5D=129", 59)
        Call .AddItem("&na%5B%5D=133", 60)
        Call .AddItem("&na%5B%5D=136", 61)
        Call .AddItem("&na%5B%5D=145", 62)
        Call .AddItem("&na%5B%5D=161", 63)
        Call .AddItem("&na%5B%5D=167", 64)
        Call .AddItem("&na%5B%5D=182", 65)
        Call .AddItem("&na%5B%5D=190", 66)
        Call .AddItem("&na%5B%5D=198", 67)
        Let .ListIndex = frm.txCountry.ListIndex
    End With

    On Error GoTo 0

Finish:
    On Error Resume Next
    Set cgf = Nothing
    Set frm = Nothing
    Exit Sub

ErrHandler:
    Call cgf.DisplayErrMessage("clsForm", "LoadCountriesHidden")
    GoTo Finish

End Sub


Public Sub SearchTeamsList()
    
    Set oClsSFL = New clsSoFifaLeagues
    Call oClsSFL.SearchTeamsList
    Set oClsSFL = Nothing
    
End Sub


Public Sub SearchPlayersList()

    Set oClsSFL = New clsSoFifaLeagues
    Call oClsSFL.SearchPlayersList
    Set oClsSFL = Nothing
    
End Sub


Public Sub GameChange()
    
    Set frm = frSoFifa
    Let frm.txGameHidden.ListIndex = frm.txGame.ListIndex
    Call LoadCountries
    Call SearchVersionList
    Call SearchTeamsList
    Call SearchPlayersList
    Set frm = Nothing
    
End Sub


Public Sub VersionChange()
    
    Set frm = frSoFifa
    Let frm.txVersionHidden.ListIndex = frm.txVersion.ListIndex
    Call SearchTeamsList
    Call SearchPlayersList
    Set frm = Nothing
    
End Sub


Public Sub CountryChange()
    
    Set frm = frSoFifa
    Let frm.txCountryHidden.ListIndex = frm.txCountry.ListIndex
    Call SearchTeamsList
    Call SearchPlayersList
    Set frm = Nothing
    
End Sub


Public Sub TeamChange()
    
    Set frm = frSoFifa
    Let frm.txTeamHidden.ListIndex = frm.txTeam.ListIndex
    Call SearchPlayersList
    Set frm = Nothing
    
End Sub


Public Sub PlayerChange()
    
    Set frm = frSoFifa
    Let frm.txPlayerHidden.ListIndex = frm.listPlayer.ListIndex
    Set frm = Nothing
    
End Sub


Public Sub BtnSearchClick()
    
    Set oClsSF = New clsSoFifa
    Call oClsSF.SearchPlayerStats
    Call BtnCopyClick
    Set oClsSF = Nothing

End Sub


Public Sub SearchGameList()

    Set oClsSFL = New clsSoFifaLeagues
    Call oClsSFL.SearchGameList
    Set oClsSFL = Nothing
    
End Sub


Public Sub SearchVersionList()

    Set oClsSFL = New clsSoFifaLeagues
    Call oClsSFL.SearchVersionList
    Set oClsSFL = Nothing
    
End Sub
