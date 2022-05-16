VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frSoFifa 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SoFifaPlayersScraper By Edgar Santa Rosa"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12675
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmSoFifa.frx":0000
   LinkTopic       =   "frSoFifa"
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   845
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton txOrderClub 
      Caption         =   "Club"
      Height          =   495
      Left            =   4560
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton txOrderOVR 
      Caption         =   "OVR"
      Height          =   495
      Left            =   4560
      TabIndex        =   26
      Top             =   2640
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.ComboBox txType 
      Height          =   375
      ItemData        =   "frmSoFifa.frx":1084A
      Left            =   1800
      List            =   "frmSoFifa.frx":1084C
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
   End
   Begin VB.ListBox listPlayer 
      Height          =   6180
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   7095
   End
   Begin MSComctlLib.Slider txMinOver 
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   3
      Min             =   60
      Max             =   99
      SelStart        =   70
      TickStyle       =   1
      TickFrequency   =   3
      Value           =   70
   End
   Begin VB.ComboBox txVersion 
      Height          =   375
      ItemData        =   "frmSoFifa.frx":1084E
      Left            =   1800
      List            =   "frmSoFifa.frx":10850
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin VB.ComboBox txGame 
      Height          =   375
      ItemData        =   "frmSoFifa.frx":10852
      Left            =   1800
      List            =   "frmSoFifa.frx":10854
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton btClear 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Picture         =   "frmSoFifa.frx":10856
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox txTeam 
      Height          =   375
      ItemData        =   "frmSoFifa.frx":13862
      Left            =   1800
      List            =   "frmSoFifa.frx":13864
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ComboBox txCountry 
      Height          =   375
      ItemData        =   "frmSoFifa.frx":13866
      Left            =   1800
      List            =   "frmSoFifa.frx":13868
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin MSComctlLib.StatusBar btStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   6360
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   661
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   22304
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btSearch 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Picture         =   "frmSoFifa.frx":1386A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton btCopy 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4560
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSoFifa.frx":16959
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txStatsToCopy 
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3960
      Width           =   5175
   End
   Begin MSComctlLib.Slider txMaxOver 
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      LargeChange     =   3
      Min             =   60
      Max             =   99
      SelStart        =   99
      TickStyle       =   1
      TickFrequency   =   3
      Value           =   99
   End
   Begin VB.ComboBox txTeamHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      ItemData        =   "frmSoFifa.frx":18996
      Left            =   5400
      List            =   "frmSoFifa.frx":18998
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.ComboBox txPlayerHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      ItemData        =   "frmSoFifa.frx":1899A
      Left            =   5400
      List            =   "frmSoFifa.frx":1899C
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.ComboBox txGameHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      ItemData        =   "frmSoFifa.frx":1899E
      Left            =   5400
      List            =   "frmSoFifa.frx":189A0
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.ComboBox txVersionHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      ItemData        =   "frmSoFifa.frx":189A2
      Left            =   5400
      List            =   "frmSoFifa.frx":189A4
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4440
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.ComboBox txCountryHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   375
      ItemData        =   "frmSoFifa.frx":189A6
      Left            =   5400
      List            =   "frmSoFifa.frx":189A8
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lbType 
      Caption         =   "Search Type"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lbThanks 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Caption         =   "THANKS TO PES6.ES AND SOFIFA.COM"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Label lbMaxOver 
      Caption         =   "Max. Overall"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lbMinOver 
      Caption         =   "Min. Overall"
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lbVersion 
      Caption         =   "Version"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lbGame 
      Caption         =   "Game"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lbTeam 
      Caption         =   "Clubs"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lbCountry 
      Caption         =   "Competition"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frSoFifa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btClear_Click()
    Call modForm.BtnClearClick
End Sub


Private Sub btCopy_Click()
    Call modForm.BtnCopyClick
End Sub


Private Sub btSearch_Click()
    Call modForm.BtnSearchClick
End Sub


Private Sub Form_Activate()
    Call modForm.FormActivate
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Let Cancel = modForm.FormClose
End Sub


Private Sub txGame_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13: Call modForm.GameChange
    End Select
End Sub

Private Sub txType_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13: Call modForm.CountryChange
    End Select
End Sub

Private Sub txVersion_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13: Call modForm.VersionChange
    End Select
End Sub


Private Sub txCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13: Call modForm.CountryChange
    End Select
End Sub


Private Sub txTeam_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case 9, 13: Call modForm.TeamChange
    End Select
End Sub


Private Sub listPlayer_Click()
    Call modForm.PlayerChange
    Call modForm.ListboxClick
End Sub
