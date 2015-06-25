VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Step 4 of 4"
   ClientHeight    =   8790
   ClientLeft      =   -105
   ClientTop       =   285
   ClientWidth     =   14790
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Previous"
      Height          =   495
      Left            =   6240
      TabIndex        =   26
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   8280
      TabIndex        =   25
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "(B) Assessment of Personal Attributes and Functional Competency (Weightage : 50%) ------------------------- (Part 2 of 2)"
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
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   3960
         Width           =   14295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Generate Result"
         Height          =   495
         Left            =   6720
         TabIndex        =   27
         Top             =   3360
         Width           =   1455
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
         Left            =   7440
         TabIndex        =   23
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate"
         Height          =   495
         Left            =   8400
         TabIndex        =   22
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "(vii) Leadership qualities"
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
         TabIndex        =   1
         Top             =   1320
         Width           =   14295
         Begin VB.CheckBox Check1 
            Caption         =   "NA"
            Height          =   375
            Left            =   6960
            TabIndex        =   7
            Top             =   600
            Width           =   615
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   4
            Left            =   10440
            TabIndex        =   6
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   3
            Left            =   8760
            TabIndex        =   5
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   2
            Left            =   7080
            TabIndex        =   4
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   1
            Left            =   5400
            TabIndex        =   3
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Height          =   375
            Index           =   0
            Left            =   3720
            TabIndex        =   2
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Unlikely to become a leader"
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
            TabIndex        =   9
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Has potential to become an excellent leader"
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
            TabIndex        =   8
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Label Label23 
         Caption         =   "Overall Grading on 'Work Output' [Total (i to vii)/7] :"
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
         TabIndex        =   24
         Top             =   2760
         Width           =   5295
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   17
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
         Height          =   375
         Left            =   9840
         TabIndex        =   16
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
         TabIndex        =   15
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
         Height          =   375
         Left            =   6480
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Public FSys As New FileSystemObject
Dim OverallGrading As Variant

Private Sub Command1_Click()
    If Check1.Value = Unchecked Then
        For i = 0 To 4
            If Option1(i).Value = True Then
                Form3.Number2 = Form3.Number2 + 1
                Form3.Point7 = Val(Label8(i).Caption)
                Form3.Score2 = Form3.Score2 + Form3.Point7
            End If
        Next
    End If
    
    Form3.TotalScore2 = Form3.Score2
    Text1.Text = Round((Form3.Score2 / Form3.Number2) * 100) / 100
    Form3.Score2 = 0
    Form3.Number2 = 0
    OverallGrading = (Form2.TotalScore1 + Form3.TotalScore2) / 2
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    Form4.Hide
    Form3.Show
End Sub

Private Sub Command4_Click()
    Dim OutStream As TextStream
    Dim InStream As TextStream
    Dim ResultFile As String
    
    ResultFile = "C:\" & CStr(Form1.EmpName) & CStr(Form1.EmpPeriod) & ".txt"
    Set OutStream = FSys.CreateTextFile(ResultFile, True, False)
    
    OutStream.WriteLine ("Assessment of " & Form1.EmpName & " (" & Form1.EmpPeriod & ")")
    OutStream.WriteBlankLines (1)
    
    OutStream.WriteLine ("                      CRITERION                        : POINTS")
    OutStream.WriteLine ("(A) ASSESSMENT OF WORK OUTPUT (WEIGHTAGE : 50%)")
    OutStream.WriteLine (" 1. Quality of work and productivity                   : " & Form2.Points1)
    OutStream.WriteLine (" 2. Work Output                                        : " & Form2.Points2)
    OutStream.WriteLine (" 3. Upkeep of equipment/area                           : " & Form2.Points3)
    OutStream.WriteLine (" 4. Initiative and drive                               : " & Form2.Points4)
    OutStream.WriteLine (" 5. Attitude to safety                                 : " & Form2.Points5)
    OutStream.WriteLine ("    OVERALL GRADING ON 'WORK OUTPUT' : " & Form2.TotalScore1)
    OutStream.WriteBlankLines (1)
    
    OutStream.WriteLine ("(B) ASSESSMENT OF PERSONAL ATTRIBUTES AND FUNCTIONAL COMPETENCY (WEIGHTAGE : 50%)")
    OutStream.WriteLine (" 1. Attendance and punctuality                         : " & Form3.Point1)
    OutStream.WriteLine (" 2. Maintenance of discipline                          : " & Form3.Point2)
    OutStream.WriteLine (" 3. Inter-personal relations                           : " & Form3.Point3)
    OutStream.WriteLine (" 4. Team spirit                                        : " & Form3.Point4)
    OutStream.WriteLine (" 5. Knowledge of procedures in the area of functioning : " & Form3.Point5)
    OutStream.WriteLine (" 6. Technical knowledge                                : " & Form3.Point6)
    OutStream.WriteLine (" 7. Leadership qualities                               : " & Form3.Point7)
    OutStream.WriteLine ("    OVERALL GRADING ON 'PERSONAL ATTRIBUTES AND FUNCTIONAL COMPETENCY' : " & Form3.TotalScore2)
    OutStream.WriteBlankLines (2)
    
    OutStream.WriteLine (" OVERALL GRADING ON 'WORK OUTPUT' : " & Form2.TotalScore1)
    OutStream.WriteLine (" OVERALL GRADING ON 'PERSONAL ATTRIBUTES AND FUNCTIONAL COMPETENCY' : " & Form3.TotalScore2)
    OutStream.WriteBlankLines (1)
    
    OutStream.WriteLine (" OVERALL GRADING : " & OverallGrading)
    
    Text2.Text = ""
    Set InStream = FSys.OpenTextFile(ResultFile, ForReading, False, TristateFalse)
    Text2.Text = InStream.ReadAll
End Sub
