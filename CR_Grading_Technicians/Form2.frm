VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Step 2 of 4"
   ClientHeight    =   8790
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   14790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Previous"
      Height          =   495
      Left            =   5280
      TabIndex        =   63
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   495
      Left            =   6840
      TabIndex        =   61
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8400
      TabIndex        =   60
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "(A) Assessment of Work Output (Weightage : 50%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton Command3 
         Caption         =   "Calculate"
         Height          =   495
         Left            =   8280
         TabIndex        =   62
         Top             =   7080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   59
         Top             =   7200
         Width           =   735
      End
      Begin VB.Frame Frame6 
         Caption         =   "(v) Attitude to safety"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   49
         Top             =   5760
         Width           =   14295
         Begin VB.CheckBox Check5 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   55
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   54
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   53
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   52
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   51
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   50
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "Observance of safety rules meticulously"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Caption         =   "Negligent towards safety"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            TabIndex        =   56
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "(iv) Initiative and drive"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   40
         Top             =   4680
         Width           =   14295
         Begin VB.CheckBox Check4 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   46
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   45
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   44
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   43
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   42
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "Takes initiative/drive to improve his/her skills/trade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "Makes no effort to improve his/her skills/trade"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            TabIndex        =   47
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "(iii) Upkeepment of equipment/area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   31
         Top             =   3600
         Width           =   14295
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   34
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   33
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check3 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   32
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Does not attend to plant or area"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            TabIndex        =   39
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Maintains cleanliness and carried out all check-outs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "(ii) Work output"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   21
         Top             =   2520
         Width           =   14295
         Begin VB.CheckBox Check2 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   30
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   26
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   23
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "Completes assigned job in time and with minimum supervision"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Extremely slow and needs constant supervision"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            TabIndex        =   27
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "(i) Quality of work and productivity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   14295
         Begin VB.CheckBox Check1 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   29
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   20
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   16
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Poor quality of work and very low product output"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11160
            TabIndex        =   15
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Excellent quality of work and highly productive"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Label Label23 
         Caption         =   "Overall Grading on 'Work Output' [Total (i to v)/5] :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   58
         Top             =   7200
         Width           =   5175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   9840
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   8160
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   6480
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Y Applies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9840
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "X Applies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Tendency to Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Tendency to X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4800
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11400
         TabIndex        =   2
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Points1 As Variant
Public Points2 As Variant
Public Points3 As Variant
Public Points4 As Variant
Public Points5 As Variant
Public TotalScore1 As Variant
Public Score1 As Variant
Public Number1 As Integer

Dim i As Integer

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    Form2.Hide
    Form3.Show
End Sub

Private Sub Command3_Click()
    If Check1.Value = Unchecked Then
        For i = 0 To 4
            If Option1(i).Value = True Then
                Number1 = Number1 + 1
                Points1 = Val(Label8(i).Caption)
                Score1 = Score1 + Points1
            End If
        Next
    End If
    
    If Check2.Value = Unchecked Then
        For i = 0 To 4
            If Option2(i).Value = True Then
                Number1 = Number1 + 1
                Points2 = Val(Label8(i).Caption)
                Score1 = Score1 + Points2
            End If
        Next
    End If
    
    If Check3.Value = Unchecked Then
        For i = 0 To 4
            If Option3(i).Value = True Then
                Number1 = Number1 + 1
                Points3 = Val(Label8(i).Caption)
                Score1 = Score1 + Points3
            End If
        Next
    End If
    
    If Check4.Value = Unchecked Then
        For i = 0 To 4
            If Option4(i).Value = True Then
                Number1 = Number1 + 1
                Points4 = Val(Label8(i).Caption)
                Score1 = Score1 + Points4
            End If
        Next
    End If
    
    If Check5.Value = Unchecked Then
        For i = 0 To 4
            If Option5(i).Value = True Then
                Number1 = Number1 + 1
                Points5 = Val(Label8(i).Caption)
                Score1 = Score1 + Points5
            End If
        Next
    End If
    TotalScore1 = Score1
    Text1.Text = Round((Score1 / Number1) * 100) / 100
    Score1 = 0
    Number1 = 0
End Sub

Private Sub Command4_Click()
    Form2.Hide
    Form1.Show
End Sub

Private Sub Form_Load()
    With Form2
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    Number1 = 0
    Score1 = 0
    TotalScore1 = 0
    Points1 = "NA"
    Points2 = "NA"
    Points3 = "NA"
    Points4 = "NA"
    Points5 = "NA"
End Sub
