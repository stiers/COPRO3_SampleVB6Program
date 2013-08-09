VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   2925
   ClientTop       =   2490
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   6120
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   450
      Left            =   255
      TabIndex        =   14
      Top             =   5160
      Width           =   2550
   End
   Begin VB.Frame Frame2 
      Height          =   2100
      Left            =   120
      TabIndex        =   5
      Top             =   2835
      Width           =   5895
      Begin VB.TextBox txtStudentAge 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   9
         Top             =   1560
         Width           =   2580
      End
      Begin VB.ComboBox cboStudentGender 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmMain.frx":0000
         Left            =   3120
         List            =   "frmMain.frx":000A
         TabIndex        =   8
         Top             =   600
         Width           =   2580
      End
      Begin VB.TextBox txtStudentName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   2580
      End
      Begin VB.TextBox txtStudentID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   2580
      End
      Begin VB.Label lblStudentName 
         Caption         =   "Student Name"
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   1185
         Width           =   1395
      End
      Begin VB.Label lblStudentGender 
         Caption         =   "Gender"
         Height          =   360
         Left            =   3120
         TabIndex        =   12
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label lblStudentAge 
         Caption         =   "Age"
         Height          =   360
         Left            =   3120
         TabIndex        =   11
         Top             =   1185
         Width           =   1395
      End
      Begin VB.Label lblStudentID 
         Caption         =   "Student ID"
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid grdMain 
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   4048
         _Version        =   393216
         BackColorBkg    =   16777215
         Appearance      =   0
      End
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "Update"
      Height          =   450
      Left            =   3255
      TabIndex        =   2
      Top             =   5160
      Width           =   2550
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   450
      Left            =   3255
      TabIndex        =   1
      Top             =   5760
      Width           =   2550
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   450
      Left            =   255
      TabIndex        =   0
      Top             =   5760
      Width           =   2550
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()

    End

End Sub

Private Sub btnSave_Click()

    Add txtStudentID, txtStudentName, cboStudentGender, txtStudentAge
    Display grdMain
    Clear

End Sub

Private Sub Form_Load()

    Display grdMain
    ColorGridRow

End Sub

Private Sub grdMain_Click()

    With grdMain
    
        For x = 0 To .Cols - 1
            .Col = x
            .CellBackColor = &HC0C0FF
        Next x
            
    End With

End Sub

Private Sub grdMain_DblClick()

    With grdMain
        txtStudentID = .TextMatrix(.Row, 1)
        txtStudentName = .TextMatrix(.Row, 2)
        cboStudentGender.Text = .TextMatrix(.Row, 3)
        txtStudentAge = .TextMatrix(.Row, 4)
    End With

End Sub

Private Sub txtStudentAge_KeyPress(i As Integer)

    Select Case i
        Case 8, 48 To 57
            i = i
            
        Case Else
            i = 0
    End Select

End Sub

Sub Clear()

    Dim Control
    
    For Each Control In Me
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
    
End Sub

Sub ColorGridRow()

    With grdMain
    
        For i = 1 To .Rows - 1
            For Columns = 1 To .Cols - 1
                If (i Mod 2) = 0 Then
             
                    .Row = i
                    .Col = Columns
                    .CellBackColor = &HC0FFC0

                End If
            Next Columns
        Next i
    End With

End Sub
