VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Step 3 of 4"
   ClientHeight    =   8790
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   14790
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Previous"
      Height          =   495
      Left            =   5280
      TabIndex        =   69
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8520
      TabIndex        =   14
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      Height          =   495
      Left            =   6840
      TabIndex        =   13
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "(B) Assessment of Personal Attributes and Functional Competency (Weightage : 50%) ------------------------- (Part 1 of 2)"
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
      Begin VB.Frame Frame7 
         Caption         =   "(vi) Technical knowledge"
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
         TabIndex        =   60
         Top             =   6720
         Width           =   14295
         Begin VB.CheckBox Check6 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   66
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option6 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   65
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option6 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   64
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option6 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   63
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option6 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   62
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option6 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   61
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Caption         =   "Exceptionally through and up-to-date technical knowledge"
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
            TabIndex        =   68
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Caption         =   "Restricted or superficial knowledge"
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
            TabIndex        =   67
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "(i) Attendance and punctuality"
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
         TabIndex        =   51
         Top             =   1320
         Width           =   14295
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   57
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   56
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   55
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   54
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   53
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   52
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Very regular and punctual"
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
            TabIndex        =   59
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Highly irregular"
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
            TabIndex        =   58
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "(ii) Maintenance of discipline"
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
         TabIndex        =   42
         Top             =   2400
         Width           =   14295
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   48
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   47
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   46
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   45
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   44
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   43
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Not amenable to discipline"
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
            TabIndex        =   50
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "Highly disciplined"
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
            TabIndex        =   49
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "(iii) Inter-personal relations"
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
         TabIndex        =   33
         Top             =   3480
         Width           =   14295
         Begin VB.CheckBox Check3 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   39
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   38
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   37
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   36
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option3 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   34
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Co-operative and cordial"
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
            TabIndex        =   41
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Uncooperative and quarrelsome"
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
            TabIndex        =   40
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "(iv) Team spirit"
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
         TabIndex        =   24
         Top             =   4560
         Width           =   14295
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   30
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   29
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   28
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   27
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option4 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   26
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check4 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   25
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "Cannot work in team"
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
            TabIndex        =   32
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "Excellent team person"
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
            TabIndex        =   31
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "(v) Knowledge of procedures in the area of functioning"
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
         TabIndex        =   15
         Top             =   5640
         Width           =   14295
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   21
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   20
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option5 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   17
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox Check5 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   16
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Caption         =   "Negligent towards procedure"
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
            TabIndex        =   23
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "Through knowledge"
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
            TabIndex        =   22
            Top             =   360
            Width           =   3015
         End
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
         TabIndex        =   12
         Top             =   720
         Width           =   2895
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
         TabIndex        =   11
         Top             =   720
         Width           =   2895
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
         TabIndex        =   10
         Top             =   360
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
         TabIndex        =   9
         Top             =   360
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
         Height          =   375
         Left            =   6480
         TabIndex        =   8
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
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   480
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
         Height          =   375
         Left            =   9840
         TabIndex        =   6
         Top             =   480
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
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   5
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
         Height          =   255
         Index           =   1
         Left            =   4800
         TabIndex        =   4
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
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   3
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
         Height          =   255
         Index           =   3
         Left            =   8160
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
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
         TabIndex        =   1
         Top             =   1080
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Point1 As Variant
Public Point2 As Variant
Public Point3 As Variant
Public Point4 As Variant
Public Point5 As Variant
Public Point6 As Variant
Public Point7 As Variant
Public TotalScore2 As Variant
Public Score2 As Variant
Public Number2 As Integer

Dim i As Integer

Private Sub Command1_Click()
    End
End Sub

Private Sub Command2_Click()
    If Check1.Value = Unchecked Then
        For i = 0 To 4
            If Option1(i).Value = True Then
                Number2 = Number2 + 1
                Point1 = Val(Label8(i).Caption)
                Score2 = Score2 + Point1
            End If
        Next
    End If
    
    If Check2.Value = Unchecked Then
        For i = 0 To 4
            If Option2(i).Value = True Then
                Number2 = Number2 + 1
                Point2 = Val(Label8(i).Caption)
                Score2 = Score2 + Point2
            End If
        Next
    End If
    
    If Check3.Value = Unchecked Then
        For i = 0 To 4
            If Option3(i).Value = True Then
                Number2 = Number2 + 1
                Point3 = Val(Label8(i).Caption)
                Score2 = Score2 + Point3
            End If
        Next
    End If
    
    If Check4.Value = Unchecked Then
        For i = 0 To 4
            If Option4(i).Value = True Then
                Number2 = Number2 + 1
                Point4 = Val(Label8(i).Caption)
                Score2 = Score2 + Point4
            End If
        Next
    End If
    
    If Check5.Value = Unchecked Then
        For i = 0 To 4
            If Option5(i).Value = True Then
                Number2 = Number2 + 1
                Point5 = Val(Label8(i).Caption)
                Score2 = Score2 + Point5
            End If
        Next
    End If
    
    If Check6.Value = Unchecked Then
        For i = 0 To 4
            If Option6(i).Value = True Then
                Number2 = Number2 + 1
                Point6 = Val(Label8(i).Caption)
                Score2 = Score2 + Point6
            End If
        Next
    End If
    Form3.Hide
    Form4.Show
End Sub

Private Sub Command3_Click()
    Form3.Hide
    Form2.Show
End Sub

Private Sub Form_Load()
    With Form3
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With
    Point1 = "NA"
    Point2 = "NA"
    Point3 = "NA"
    Point4 = "NA"
    Point5 = "NA"
    Point6 = "NA"
    Point7 = "NA"
    
    Score2 = 0
    Number2 = 0
End Sub

