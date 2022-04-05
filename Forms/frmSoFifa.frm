VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frSoFifa 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SoFifa Old Stats Converter - By EdgarSantaRosa"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000004&
   Icon            =   "frmSoFifa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   611
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txPlayerHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   405
      ItemData        =   "frmSoFifa.frx":1084A
      Left            =   5520
      List            =   "frmSoFifa.frx":1084C
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox txClubHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   405
      ItemData        =   "frmSoFifa.frx":1084E
      Left            =   5520
      List            =   "frmSoFifa.frx":10850
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ComboBox txCountryHidden 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   405
      ItemData        =   "frmSoFifa.frx":10852
      Left            =   5520
      List            =   "frmSoFifa.frx":10854
      Locked          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox txPlayer 
      Height          =   405
      ItemData        =   "frmSoFifa.frx":10856
      Left            =   1800
      List            =   "frmSoFifa.frx":10858
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox txClub 
      Height          =   405
      ItemData        =   "frmSoFifa.frx":1085A
      Left            =   1800
      List            =   "frmSoFifa.frx":1085C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
   Begin VB.ComboBox txCountry 
      Height          =   405
      ItemData        =   "frmSoFifa.frx":1085E
      Left            =   1800
      List            =   "frmSoFifa.frx":10860
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
   Begin MSComctlLib.StatusBar btStatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   8790
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   661
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9313
            Text            =   "READY"
            TextSave        =   "READY"
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
      Height          =   1095
      Left            =   4080
      Picture         =   "frmSoFifa.frx":10862
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton btClear 
      BackColor       =   &H8000000E&
      Height          =   1095
      Left            =   2880
      Picture         =   "frmSoFifa.frx":13951
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1095
   End
   Begin VB.ComboBox txVersion 
      Height          =   405
      ItemData        =   "frmSoFifa.frx":1695D
      Left            =   1800
      List            =   "frmSoFifa.frx":1695F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton btCopy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy To Clipboard"
      Height          =   1095
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txStatsToCopy 
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3720
      Width           =   5055
   End
   Begin VB.PictureBox imgPes6ES 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Height          =   615
      Left            =   0
      Picture         =   "frmSoFifa.frx":16961
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3240
      Width           =   5415
   End
   Begin VB.Label lbPlayer 
      Caption         =   "Player"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1590
      Width           =   1575
   End
   Begin VB.Label lbClub 
      Caption         =   "Clubs"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1110
      Width           =   1575
   End
   Begin VB.Label lbCountry 
      Caption         =   "Competition"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   630
      Width           =   1575
   End
   Begin VB.Label lbVersion 
      Caption         =   "Version"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   150
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


Private Sub txClub_Validate(Cancel As Boolean)
    Call modForm.ClubChange
End Sub


Private Sub txCountry_Validate(Cancel As Boolean)
    Call modForm.CountryChange
End Sub


Private Sub txPlayer_Validate(Cancel As Boolean)
    Call modForm.PlayerChange
End Sub


Private Sub txVersion_Validate(Cancel As Boolean)
    Call modForm.VersionChange
End Sub
