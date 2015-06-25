VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Step 1 of 4"
   ClientHeight    =   8790
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh This Page"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continue"
      Height          =   495
      Left            =   7080
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Form1.frx":0000
      Left            =   6000
      List            =   "Form1.frx":001F
      TabIndex        =   3
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6000
      TabIndex        =   1
      Top             =   1080
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "S.K. Jha, Head, SFDS, AFD, BARC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4920
      TabIndex        =   8
      Top             =   6240
      Width           =   7455
   End
   Begin VB.Label Label3 
      Caption         =   "Developed for"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4920
      TabIndex        =   7
      Top             =   5760
      Width           =   7095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Period :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4800
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EmpName As String
Public EmpPeriod As String

Private Sub Command1_Click()
    EmpName = Text1.Text
    EmpPeriod = Combo1.Text
    If EmpName = "" Or EmpPeriod = "" Then
        MsgBox ("You must enter all details.")
    Else
        MsgBox ("Continue for " + EmpName + " during " + EmpPeriod + "?")
        Form1.Hide
        Form2.Show
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    Combo1.Text = ""
End Sub

Private Sub Form_Load()
    With Form1
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    EmpName = ""
    EmpPeriod = ""
End Sub
