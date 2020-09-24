VERSION 5.00
Begin VB.Form frmRandomNumbers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Random Numbers"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmRandomNumbers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Random Numbers"
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
      Picture         =   "frmRandomNumbers.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Random Numbers"
      Height          =   4095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear List Box"
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
         Left            =   3840
         Picture         =   "frmRandomNumbers.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   2175
      End
      Begin VB.CommandButton cmdRI 
         Caption         =   "Random &Integers"
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
         Left            =   3840
         Picture         =   "frmRandomNumbers.frx":15D6
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdRN 
         Caption         =   "Random &Decimals"
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
         Left            =   3840
         Picture         =   "frmRandomNumbers.frx":1EA0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   2175
      End
      Begin VB.ListBox lstNumbers 
         Height          =   3570
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Height          =   420
         Left            =   4200
         TabIndex        =   5
         Top             =   3600
         Width           =   1305
      End
   End
End
Attribute VB_Name = "frmRandomNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
'-Clear the listbox
lstNumbers.Clear
End Sub

Private Sub cmdDone_Click()
'-Unload frmRandomNumbers
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdRI_Click()
Randomize   '-Randomize the random numbers, so diffrent every time
'-Declare varibles
Dim X As Integer
'-Setup loop to make 50 random integer numbers
For X = 1 To 50
    '-Add 50 random integer numbers
    lstNumbers.AddItem Int(Rnd() * 100 + 1)
    '-Int(Rnd() = random integer number and a
    '-math expression to set the scope of the
    '-random integer, in this case it will make
    '-numbers between 1-100
    If X = 50 Then  '-If x = 50 the loop stops
        Exit For
    End If
'-Continue the loop
Next X
End Sub

Private Sub cmdRN_Click()
Randomize   '-Randomize the random numbers, so diffrent every time
'-Declare varibles
Dim X As Integer
'-Setup loop to make 50 random numbers
For X = 1 To 50
    '-Add 50 random numbers
    lstNumbers.AddItem Rnd  '-Rnd = random number
    If X = 50 Then  '-If x = 50 the loop stops
        Exit For
    End If
'-Continue the loop
Next X
End Sub

Private Sub Label1_Click()
'-Setup message box
MsgBox ("Randoms Decimals or randoms numbers are numbers greater than 0 and less than 1, Random Integers are numbers greater than 1 and less than whatever you want.")
End Sub
