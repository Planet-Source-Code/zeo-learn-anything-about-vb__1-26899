VERSION 5.00
Begin VB.Form frmColors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colors"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmColors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done Learing About Colors"
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
      Left            =   0
      Picture         =   "frmColors.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   6015
   End
   Begin VB.CommandButton cmdQBColor 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "QB Colors"
      Height          =   2415
      Left            =   0
      TabIndex        =   26
      Top             =   3960
      Width           =   5895
      Begin VB.PictureBox picQBColor 
         AutoRedraw      =   -1  'True
         Height          =   735
         Left            =   240
         ScaleHeight     =   675
         ScaleWidth      =   5235
         TabIndex        =   27
         Top             =   240
         Width           =   5295
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "14"
         Height          =   495
         Index           =   14
         Left            =   4440
         TabIndex        =   18
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "13"
         Height          =   495
         Index           =   13
         Left            =   3720
         TabIndex        =   17
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "12"
         Height          =   495
         Index           =   12
         Left            =   3000
         TabIndex        =   16
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "11"
         Height          =   495
         Index           =   11
         Left            =   2280
         TabIndex        =   15
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "10"
         Height          =   495
         Index           =   10
         Left            =   1560
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "9"
         Height          =   495
         Index           =   9
         Left            =   840
         TabIndex        =   13
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "8"
         Height          =   495
         Index           =   8
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "6"
         Height          =   495
         Index           =   6
         Left            =   4440
         TabIndex        =   10
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "5"
         Height          =   495
         Index           =   5
         Left            =   3720
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "4"
         Height          =   495
         Index           =   4
         Left            =   3000
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "3"
         Height          =   495
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "2"
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "15"
         Height          =   495
         Index           =   15
         Left            =   5160
         TabIndex        =   19
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "7"
         Height          =   495
         Index           =   7
         Left            =   5160
         TabIndex        =   11
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdQBColor 
         Caption         =   "1"
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "RGB Colors"
      Height          =   3975
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   6015
      Begin VB.PictureBox picColor 
         AutoRedraw      =   -1  'True
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   480
         ScaleHeight     =   915
         ScaleWidth      =   4995
         TabIndex        =   22
         Top             =   360
         Width           =   5055
      End
      Begin VB.HScrollBar hsBlue 
         Height          =   255
         LargeChange     =   5
         Left            =   480
         Max             =   255
         TabIndex        =   2
         Top             =   3000
         Width           =   5055
      End
      Begin VB.HScrollBar hsGreen 
         Height          =   255
         LargeChange     =   5
         Left            =   480
         Max             =   255
         TabIndex        =   1
         Top             =   2400
         Width           =   5055
      End
      Begin VB.HScrollBar hsRed 
         Height          =   255
         LargeChange     =   5
         Left            =   480
         Max             =   255
         TabIndex        =   0
         Top             =   1800
         Width           =   5055
      End
      Begin VB.Label Label4 
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
         Left            =   2160
         TabIndex        =   3
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2520
         TabIndex        =   25
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Papyrus"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   1440
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
'-Unload frmColors
Unload Me
'-make frmMain visible
frmMain.Visible = True
End Sub

Private Sub cmdQBColor_Click(Index As Integer)
'-Select index of the button clicked
Select Case Index
    Case "0"    '-Black
        picQBColor.BackColor = QBColor(0)
    Case "1"    '-Blue
        picQBColor.BackColor = QBColor(1)
    Case "2"    '-Green
        picQBColor.BackColor = QBColor(2)
    Case "3"    '-Cyan
        picQBColor.BackColor = QBColor(3)
    Case "4"    '-Red
        picQBColor.BackColor = QBColor(4)
    Case "5"    '-Magenta
        picQBColor.BackColor = QBColor(5)
    Case "6"    '-Yellow
        picQBColor.BackColor = QBColor(6)
    Case "7"    '-Gray
        picQBColor.BackColor = QBColor(7)
    Case "8"    '- Dark Gray
        picQBColor.BackColor = QBColor(8)
    Case "9"    '-Light Blue
        picQBColor.BackColor = QBColor(9)
    Case "10"   '-Light Green
        picQBColor.BackColor = QBColor(10)
    Case "11"   '-Light Cyan
        picQBColor.BackColor = QBColor(11)
    Case "12"   '-Light Red
        picQBColor.BackColor = QBColor(12)
    Case "13"   '-Light Magenta
        picQBColor.BackColor = QBColor(13)
    Case "14"   '-Light Yellow
        picQBColor.BackColor = QBColor(14)
    Case "15"   '-White
        picQBColor.BackColor = QBColor(15)
'-End the case select index
End Select
End Sub

Private Sub hsBlue_Change()
'-Get the value of the scroll bar
Blue = (hsBlue.Value)
'-Mix the colors using the Red Green Blue method
picColor.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub hsGreen_Change()
'-Get the value of the scroll bar
Green = (hsGreen.Value)
'-Mix the colors using the Red Green Blue method
picColor.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub hsRed_Change()
'-Get the value of the scroll bar
Red = (hsRed.Value)
'-Mix the colors using the Red Green Blue method
picColor.BackColor = RGB(Red, Green, Blue)
End Sub

Private Sub Label4_Click()
'-Setup message box
MsgBox ("The RGB method mixes the amount or value of each color together to create billions of diffrent colors.")
End Sub
