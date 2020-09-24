VERSION 5.00
Begin VB.Form frmListBoxes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Boxes"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmListBoxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Done Learning About ListBoxes"
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
      Picture         =   "frmListBoxes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Loading List Boxes"
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5895
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
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
         Picture         =   "frmListBoxes.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin VB.ListBox lstLoad 
         Height          =   2205
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3495
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
         Left            =   4080
         TabIndex        =   8
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Adding and Saving Items"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save List Box"
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
         Picture         =   "frmListBoxes.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Item"
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
         Left            =   3840
         Picture         =   "frmListBoxes.frx":2C20
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.ListBox lstAdd 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmListBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
'-Declare varibles
Dim Add As String
'-Gather information from the user
Add = InputBox("Enter the item to be added", "Add Item...", "Hello")
'-Put the information from Add into the list box
lstAdd.AddItem Add
End Sub

Private Sub cmdLoad_Click()
'-Declare varibles
Dim directory$
Dim MyString$
'-If it messes up
 On Error GoTo errhandle
    '-Clear the list
    lstLoad.Clear
    '-Find the fill to open
    directory$ = App.Path & "\Saves\Listbox.txt"
    '-Load the file
    Open directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        lstLoad.AddItem (MyString$)
    Wend
    '-Close the file
    Close #1
errhandle:
End Sub

Private Sub cmdSave_Click()
'-Declare Varibles
Dim Savelist As Long
Dim directory$
'-If it messes up it will not do anything
On Error Resume Next
'-Place to save
directory$ = App.Path & "\Saves\Listbox.txt"
'-Write File
Open directory$ For Output As #1
    For Savelist& = 0 To lstAdd.ListCount - 1
        Print #1, (lstAdd.List(Savelist&))
    Next Savelist&
    '-Close file
    Close #1
'-Setup message box
MsgBox ("The list was saved to " & App.Path & "\Saves\Listbox.txt")
End Sub

Private Sub Command1_Click()
'-Unload frmListboxes
Unload Me
'-make frmMain visible
frmMain.Visible = True
End Sub

Private Sub Label1_Click()
'-Setup message box
MsgBox ("Click either list box to clear it.")
End Sub

Private Sub lstAdd_Click()
'-Clear lstAdd
lstAdd.Clear
End Sub

Private Sub lstLoad_Click()
'-Clear lstLoad
lstLoad.Clear
End Sub
