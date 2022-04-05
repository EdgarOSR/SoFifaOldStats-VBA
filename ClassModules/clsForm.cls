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

Private oForm As frSoFifa
Private oClsGF As clsGlobalFunc
Private oClsSF As clsSoFifa
Private oClsSFL As clsSoFifaLeagues


Public Function FormClose() As Integer

    Let FormClose = False

    Set oClsGF = New clsGlobalFunc
    Set oForm = frSoFifa
    
    On Error GoTo ErrHandler
    
    Select Case (MsgBox("Do you really like to exit?", vbYesNo, "Exit"))
        Case vbNo
            Let FormClose = True
            GoTo Finish
        Case Else
    End Select

    Call oForm.Hide
    Call Unload(oForm)

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set oForm = Nothing
    Exit Function
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "FormClose")
    GoTo Finish
    
End Function


Private Sub LoadVersions()
    
    Set oClsGF = New clsGlobalFunc
    Set oForm = frSoFifa

    With oForm.txVersion
        Call .Clear
        Call .AddItem("070001", 0)
        Call .AddItem("070002", 1)
        Call .AddItem("080001", 2)
        Call .AddItem("080002", 3)
        Call .AddItem("090001", 4)
        Call .AddItem("090002", 5)
        Let .ListIndex = 0
    End With

    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set oForm = Nothing
    Exit Sub
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "LoadVersions")
    GoTo Finish
 
End Sub


Private Sub DisplayDependencies()

    Dim vMsg As String
    
    Set oClsGF = New clsGlobalFunc
    
    On Error GoTo ErrHandler
    
    Let vMsg = vbNullString
    Let vMsg = vMsg & "You must install for correct functionality:" & vbCrLf
    Let vMsg = vMsg & "* SELENIUM VBA: https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0" & vbCrLf
    Let vMsg = vMsg & "* CHROME DRIVER: https://chromedriver.storage.googleapis.com/index.html" & vbCrLf & vbCrLf
    Let vMsg = vMsg & "Put CHROME DRIVE where SELENIUM was installed, Chrome Drive must be according your Chrome version" & vbCrLf
    Let vMsg = vMsg & "Regards, Edgar Santa Rosa"

    Call MsgBox(vMsg, vbOKOnly)
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Exit Sub
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "DisplayDependencies")
    GoTo Finish

End Sub


Public Sub BtnCopyClick()
    
    Dim oDto As VB.Clipboard
    
    Set oClsGF = New clsGlobalFunc
    Set oDto = VB.Clipboard
    Set oForm = frSoFifa
    
    On Error GoTo ErrHandler
    
    With oForm
        If (.txStatsToCopy.Text = vbNullString) Then GoTo Finish
        Call oDto.Clear
        Let .txStatsToCopy.Text = Replace(.txStatsToCopy.Text, "?", "*")
        Call oDto.SetText(.txStatsToCopy.Text)
        Call MsgBox("Paste it in PES6 Editor", vbInformation)
    End With
    
    On Error GoTo 0

Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set oForm = Nothing
    Set oDto = Nothing
    Exit Sub

ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "BtnCopyClick")
    GoTo Finish
    
    
End Sub


Public Sub BtnClearClick()
    
    Set oClsGF = New clsGlobalFunc
    Set oForm = frSoFifa
    
    On Error GoTo ErrHandler

    With oForm
        Call .txVersion.SetFocus
        Let .txStatsToCopy.Text = vbNullString
        Let .txVersion.ListIndex = 0
        Let .txCountry.ListIndex = 0
        Let .txCountryHidden.ListIndex = 0
    End With
    
    Call SearchClubsList
    Call SearchPlayersList
    
    On Error GoTo 0

Finish:
    On Error Resume Next
    Set oForm = Nothing
    Set oClsGF = Nothing
    Exit Sub

ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "BtnClearClick")
    GoTo Finish
    
End Sub


Public Sub FormActivate()
    
    Call DisplayDependencies
    Call LoadVersions
    Call LoadCountries
    Call LoadCountriesHidden
    Call BtnClearClick

End Sub


Private Sub LoadCountries()

    Dim vIdx As Integer

    Set oClsGF = New clsGlobalFunc
    Set oForm = frSoFifa

    On Error GoTo ErrHandler

    With oForm.txCountry
        Call .Clear
        
        Let vIdx = 0
        Call .AddItem("National Teams", vIdx)
        
        Let vIdx = vIdx + 1
        Call .AddItem("Argentina", vIdx)

        If (Val(oForm.txVersion.Text) <> 150001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Brazil", vIdx)
        End If

        If (Val(oForm.txVersion.Text) >= 140001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Chile", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("Colombia", vIdx)
        End If

        If (Val(oForm.txVersion.Text) > 200028) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Bolivia", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("Ecuador", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("Paraguay", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("Peru", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("Uruguay", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("Venezuela", vIdx)
        End If

        Let vIdx = vIdx + 1
        Call .AddItem("Mexico", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("United States", vIdx)

        Let vIdx = vIdx + 1
        Call .AddItem("Austria", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Belgium", vIdx)

        If (Val(oForm.txVersion.Text) >= 190001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Croatia", vIdx)
        End If

        Let vIdx = vIdx + 1
        Call .AddItem("Czech", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Denmark", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("England", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("France", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Germany", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Greece", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Italy", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Netherlands", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Norway", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Poland", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Portugal", vIdx)

        If (Val(oForm.txVersion.Text) >= 200001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Romania", vIdx)
        End If

        If (Val(oForm.txVersion.Text) >= 100001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Russia", vIdx)
        End If
        
        Let vIdx = vIdx + 1
        Call .AddItem("Scotland", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Sweden", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Switzerland", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("Turkey", vIdx)
        
        If (Val(oForm.txVersion.Text) >= 140001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Ukraine", vIdx)
        End If
        
        If (Val(oForm.txVersion.Text) >= 80001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Australia", vIdx)
        End If
        
        If (Val(oForm.txVersion.Text) >= 170001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("Japan", vIdx)
        End If
        
        Let vIdx = vIdx + 1
        Call .AddItem("Korea", vIdx)

        Let .ListIndex = 0
        
    End With


    
    On Error GoTo 0

Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set oForm = Nothing
    Exit Sub

ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "LoadCountries")
    GoTo Finish

End Sub


Private Sub LoadCountriesHidden()

    Dim vIdx As Integer

    Set oClsGF = New clsGlobalFunc
    Set oForm = frSoFifa

    On Error GoTo ErrHandler

    With oForm.txCountryHidden
        Call .Clear
        
        Let vIdx = 0
        Call .AddItem("type=national", vIdx)
        
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=52", vIdx)

        If (Val(oForm.txVersion.Text) <> 150001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=54", vIdx)
        End If

        If (Val(oForm.txVersion.Text) >= 140001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=55", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=56", vIdx)
        End If

        If (Val(oForm.txVersion.Text) > 200028) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=53", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=57", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=58", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=59", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=60", vIdx)
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=0&na%5B%5D=61", vIdx)
        End If

        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=1&na%5B%5D=83", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=1&na%5B%5D=95", vIdx)
        
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=4", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=7", vIdx)

        If (Val(oForm.txVersion.Text) >= 190001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=10", vIdx)
        End If
        
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=12", vIdx)
        Let vIdx = vIdx + 1
       Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=13", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=14", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=18", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=21", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=22", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=27", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=34", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=36", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=37", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=38", vIdx)

        If (Val(oForm.txVersion.Text) >= 200001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=39", vIdx)
        End If

        If (Val(oForm.txVersion.Text) >= 100001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=40", vIdx)
        End If
        
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=42", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=46", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=47", vIdx)
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=48", vIdx)
        
        If (Val(oForm.txVersion.Text) >= 140001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=2&na%5B%5D=49", vIdx)
        End If
        
        If (Val(oForm.txVersion.Text) >= 80001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=7&na%5B%5D=195", vIdx)
        End If
        
        If (Val(oForm.txVersion.Text) >= 170001) Then
            Let vIdx = vIdx + 1
            Call .AddItem("type=club&ct%5B%5D=6&na%5B%5D=163", vIdx)
        End If
        
        Let vIdx = vIdx + 1
        Call .AddItem("type=club&ct%5B%5D=6&na%5B%5D=167", vIdx)

        Let .ListIndex = oForm.txCountry.ListIndex

    End With

    On Error GoTo 0

Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set oForm = Nothing
    Exit Sub

ErrHandler:
    Call oClsGF.DisplayErrMessage("clsForm", "LoadCountriesHidden")
    GoTo Finish

End Sub


Public Sub SearchClubsList()
    
    Set oClsSFL = New clsSoFifaLeagues
    Call oClsSFL.SearchClubsList
    Set oClsSFL = Nothing
    
End Sub


Public Sub SearchPlayersList()

    Set oClsSFL = New clsSoFifaLeagues
    Call oClsSFL.SearchPlayersList
    Set oClsSFL = Nothing
    
End Sub

Public Sub VersionChange()
    
    Set oForm = frSoFifa
    
    With oForm
        Call .txCountry.Clear
        Call .txCountryHidden.Clear
        Call .txClub.Clear
        Call .txClubHidden.Clear
        Call .txPlayer.Clear
        Call .txPlayerHidden.Clear
    End With

    Call LoadCountries
    Call LoadCountriesHidden
    Call SearchClubsList
    Call SearchPlayersList
    
    Set oForm = Nothing
    
End Sub


Public Sub CountryChange()
    
    Set oForm = frSoFifa
    
    With oForm
        Let .txCountryHidden.ListIndex = oForm.txCountry.ListIndex
        Call .txClub.Clear
        Call .txClubHidden.Clear
        Call .txPlayer.Clear
        Call .txPlayerHidden.Clear
    End With

    Call SearchClubsList
    Call SearchPlayersList
    
    Set oForm = Nothing
    
End Sub


Public Sub ClubChange()
    
    Set oForm = frSoFifa
    
    With oForm
        Let .txClubHidden.ListIndex = oForm.txClub.ListIndex
        Call .txPlayer.Clear
        Call .txPlayerHidden.Clear
    End With

    Call SearchPlayersList
    
    Set oForm = Nothing
    
End Sub


Public Sub PlayerChange()
    
    Set oForm = frSoFifa
    
    With oForm
        Let .txPlayerHidden.ListIndex = oForm.txPlayer.ListIndex
    End With
    
    Set oForm = Nothing
    
End Sub


Public Sub BtnSearchClick()
    
    Set oClsSF = New clsSoFifa
    Call oClsSF.SearchPlayerStats
    Call BtnCopyClick
    Set oClsSF = Nothing

End Sub
