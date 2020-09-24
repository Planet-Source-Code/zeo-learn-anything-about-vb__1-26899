VERSION 5.00
Begin VB.Form frmKill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete Files"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmKill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learning About Deleting Files"
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
      Picture         =   "frmKill.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Delete"
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton Command2 
         Caption         =   "Delete Dum&my File"
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
         Left            =   120
         Picture         =   "frmKill.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Make Dumm&y File"
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
         Left            =   120
         Picture         =   "frmKill.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   2415
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
         Left            =   4560
         TabIndex        =   5
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblLocate 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmKill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmKill
Unload Me
'-Make frmMain visible
frmMain.Visible = True
End Sub

Private Sub Command1_Click()
'-Change the caption of the label to tell where the file is
lblLocate.Caption = "A Dummy File with the name of Dummy.txt was created in " & Chr(13) & App.Path & "\Saves\Dummy.txt"
'-Declare Varibles
Dim directory$
'-On error it will not save
On Error Resume Next
    '-Save to where my program is and in the folder "Saves"
    directory$ = App.Path & "\Saves\Dummy.txt"
    '-Writing the file
    Open directory$ For Output As #1
    Print #1, (lblLocate)
    '-Close the file
    Close #1
'-Setup msgbox
End Sub

Private Sub Command2_Click()
'-Change the caption of the label to tell what file was deleted
lblLocate.Caption = "A  File with the name of Dummy.txt was  deleted from " & Chr(13) & App.Path & "\Saves\Dummy.txt"
'-Delete the file
Kill (App.Path & "\Saves\Dummy.txt")
End Sub

Private Sub Label1_Click()
'-Setup message box
MsgBox ("To delete files you use the Kill statment," & Chr(13) & "Example" & Chr(13) & "Kill(C:\Windows\Desktop\Bob.ico)")
End Sub
