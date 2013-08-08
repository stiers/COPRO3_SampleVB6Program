VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   2655
   ClientTop       =   2640
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   6120
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   14
      Top             =   5760
      Width           =   2550
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3240
      TabIndex        =   13
      Top             =   5760
      Width           =   2550
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3240
      TabIndex        =   12
      Top             =   5160
      Width           =   2550
   End
   Begin VB.Frame Frame2 
      Height          =   2700
      Left            =   105
      TabIndex        =   10
      Top             =   120
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   4048
         _Version        =   393216
         BackColorBkg    =   16777215
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   105
      TabIndex        =   1
      Top             =   2835
      Width           =   5895
      Begin VB.TextBox Text1 
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
         TabIndex        =   5
         Top             =   600
         Width           =   2580
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   4
         Top             =   1560
         Width           =   2580
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "FRM_main.frx":0000
         Left            =   3120
         List            =   "FRM_main.frx":000A
         TabIndex        =   3
         Top             =   600
         Width           =   2580
      End
      Begin VB.TextBox Text3 
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
         TabIndex        =   2
         Top             =   1560
         Width           =   2580
      End
      Begin VB.Label Label4 
         Caption         =   "Student ID"
         Height          =   360
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Age"
         Height          =   360
         Left            =   3120
         TabIndex        =   8
         Top             =   1185
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Gender"
         Height          =   360
         Left            =   3120
         TabIndex        =   7
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "Student Name"
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   1185
         Width           =   1395
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   2550
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    mysave Text1, Text2, Combo1, Text3
    display grid1
    
    Clear

End Sub

Private Sub Command3_Click()

    End

End Sub

Private Sub Text3_KeyPress(i As Integer)

    Select Case i
        Case 8, 48 To 57
                i = i
        Case Else
                i = 0
    End Select

End Sub

Private Sub Form_Load()

    display grid1
    ColorMyRow
    
End Sub

Sub Clear()

    Dim Control
    
    For Each Control In Me
    
        If TypeOf Control Is TextBox Then
        
            Control.Text = ""
            
        End If
    
    Next Control
    
End Sub

Sub ColorMyRow()

    With grid1
    
        For i = 1 To .Rows - 1
        
            For mycol = 1 To .Cols - 1
            
                If (i Mod 2) = 0 Then
             
                    .Row = i
                    .Col = mycol
                    .CellBackColor = &HC0FFC0

                End If

            Next mycol

        Next i
    
    End With

End Sub

Private Sub grid1_Click()

    With grid1
    
        For X = 0 To .Cols - 1
            .Col = X
            .CellBackColor = &HC0C0FF
        
        Next X
            
    End With
    
End Sub

Public Sub grid1_DblClick()

    With grid1
    
        Text1 = .TextMatrix(.Row, 1)
        Text2 = .TextMatrix(.Row, 2)
        Combo1.Text = .TextMatrix(.Row, 3)
        Text3 = .TextMatrix(.Row, 4)
    
    End With

End Sub
