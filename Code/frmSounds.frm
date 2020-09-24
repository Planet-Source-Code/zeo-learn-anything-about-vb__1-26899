VERSION 5.00
Begin VB.Form frmSounds 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sounds"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmSounds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Done Learning About Sounds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      Picture         =   "frmSounds.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play Sound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "frmSounds.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-Copy and paste this to be able to play sounds-----\/
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
'----------------------------------------------------/\

Private Sub cmdPlay_Click()
'-File that you want to play \/
 soundfile$ = App.Path & "\Sounds\ywin2.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    HaHa = sndPlaySound(soundfile$, wFlags%)
End Sub

Private Sub Command1_Click()
'-Unload frmSounds
Unload Me
'-make frmMain visible
frmMain.Visible = True
End Sub
