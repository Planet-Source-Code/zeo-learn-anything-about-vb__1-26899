VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCommD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Common Dialog"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "frmCommD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Common Dialogs"
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
      Picture         =   "frmCommD.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Common Dialog Options"
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin MSComDlg.CommonDialog cdPrint 
         Left            =   3480
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPrinter 
         Caption         =   "&Printer Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Picture         =   "frmCommD.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog cdHelp 
         Left            =   3480
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Picture         =   "frmCommD.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog cdFont 
         Left            =   3480
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "&Font Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2280
         Picture         =   "frmCommD.frx":1D30
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog cdColor 
         Left            =   1440
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdOpen 
         Left            =   1440
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Open..."
         FileName        =   "New File"
         Filter          =   "Nothing (*.Nothing)"
      End
      Begin MSComDlg.CommonDialog cdSave 
         Left            =   1440
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Save..."
         FileName        =   "New File"
         Filter          =   "Nothing (*.Nothing)"
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Picture         =   "frmCommD.frx":25FA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "&Color Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Picture         =   "frmCommD.frx":2904
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         Picture         =   "frmCommD.frx":2C0E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Me!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   3600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCommD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-ALL THE CODE IS SELF EXPLANATORY WILL NOT COMMENT
'-THE cdxxxxx IS THE NAME OF THE COMMON DIALOG FOR EACH
'-COMMAND THEN THE .SHOWxxxxx TELLS WHAT TO SHOW

Private Sub cmdColor_Click()
cdColor.ShowColor
End Sub

Private Sub cmdDone_Click()
'-Unload frmCommD
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdFont_Click()
'-If it cant find any fonts it will not crash
On Error GoTo err
cdFont.ShowFont
err:
End Sub

Private Sub cmdHelp_Click()
cdHelp.ShowHelp
End Sub

Private Sub cmdPrinter_Click()
cdPrint.ShowPrinter
End Sub

Private Sub cmdSave_Click()
cdSave.ShowSave
End Sub

Private Sub cmdOpen_Click()
cdOpen.ShowOpen
End Sub

Private Sub Label1_Click()
'-Setup message box
MsgBox ("Some of the common dialogs may not work on your machine, Also they do nothing but show you how to open them and set them up.")
End Sub
