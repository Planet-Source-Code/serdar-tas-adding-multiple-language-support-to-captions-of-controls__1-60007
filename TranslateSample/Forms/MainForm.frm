VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multilanguage Sample"
   ClientHeight    =   1920
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   1740
      Left            =   2370
      TabIndex        =   5
      Top             =   60
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   3069
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   0
      BackColorSel    =   16761024
      ForeColorSel    =   0
      FocusRect       =   0
      ScrollBars      =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   ""
   End
   Begin VB.CommandButton Button2 
      Caption         =   "Command1"
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Tag             =   "3"
      Top             =   1485
      Width           =   2085
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   330
      Width           =   2100
   End
   Begin VB.CommandButton Button1 
      Caption         =   "Command1"
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   990
      Width           =   2085
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Tag             =   "2"
      Top             =   705
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Tag             =   "1"
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Button3_Click()
    GridHeaders Grid, 0, Language
End Sub

Private Sub Form_Load()
    Language = LanguageID(App.Path & "\Language.ini")
    Connect App.Path & "\Database\Sample.mdb"
    FillCombo Combo1, "Select * from Cnst_LanguageNames"
    SelectLanguage MainForm, Language
    GridHeaders Grid, 0, Language
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveLanguageID App.Path & "\Language.ini", Language
End Sub

Private Sub Button1_Click()
    SaveLanguageID App.Path & "\Language.ini", Combo1.ItemData(Combo1.ListIndex)
    Language = Combo1.ItemData(Combo1.ListIndex)
    SelectLanguage MainForm, Language
    GridHeaders Grid, 0, Language
End Sub
Private Sub Button2_Click()
    PopUp 1, Language, vbInformation
End Sub
