VERSION 5.00
Begin VB.Form frmAssignment7 
   AutoRedraw      =   -1  'True
   Caption         =   "Tic-Tac-Toe"
   ClientHeight    =   4125
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6810
   Icon            =   "A7_QuachA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDisplay 
      Caption         =   "Totals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   2175
      Begin VB.Label lblDraws 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   9
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblPlayerO 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lblPlayerX 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblGPlayed 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Draws:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Player O:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Player X:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Played:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   8
      Left            =   3000
      TabIndex        =   18
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   7
      Left            =   1560
      TabIndex        =   17
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   3000
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   1560
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   3000
      TabIndex        =   12
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblTicTacToe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   4200
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   4200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   2880
      X2              =   2880
      Y1              =   0
      Y2              =   3960
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   1440
      Y1              =   0
      Y2              =   3960
   End
   Begin VB.Label lblCopyRight 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quach Inc © 2016"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset Scores"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmAssignment7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MAX = 8
Dim GPlayed As Integer
Dim PlayerXTotal As Integer
Dim PlayerOTotal As Integer
Dim DrawTotal As Integer
Dim TicNum As Integer
Dim WinnerCount(0 To MAX) As Integer
Dim Winner As Integer



Private Sub Form_Load()
    'Intializes the variables
    Dim X As Integer
    
    GPlayed = 0
    PlayerXTotal = 0
    PlayerOTotal = 0
    DrawTotal = 0
    TicNum = 0
    Winner = 0
    
    For X = 0 To MAX
        WinnerCount(X) = X + 3
    Next X
    
    lblGPlayed.Caption = Str$(GPlayed)
    lblPlayerX.Caption = Str$(PlayerXTotal)
    lblPlayerO.Caption = Str$(PlayerOTotal)
    lblDraws.Caption = Str$(DrawTotal)

End Sub



Private Sub mnuAbout_Click()
    
    MsgBox "Created by Alex Quach, 2016.", vbQuestion, "About"
    
End Sub

Private Sub mnuExit_Click()
    'Asks the user if they would like to exit the program
    Dim DType As Integer
    Dim DTitle As String
    Dim DMSg As String
    Dim Response As Integer
    
    DType = vbYesNo + vbQuestion
    DTitle = "Exit"
    DMSg = "Do you wish to exit?"
    
    Response = MsgBox(DMSg, DType, DTitle)
    
    If Response = vbYes Then
        Beep
        End
    End If

End Sub

Private Sub mnuReset_Click()
    
    Dim DType As Integer
    Dim DTitle As String
    Dim DMSg As String
    Dim Response As Integer
    'Resets the user's scores
    DType = vbYesNo + vbQuestion
    DTitle = "Reset scores"
    DMSg = "Are you sure?  The previous scores will be permanantly erased."
    
    
    If GPlayed > 0 And PlayerXTotal > 0 And PlayerOTotal > 0 And DrawTotal > 0 Then
        Response = MsgBox(DMSg, DType, DTitle)
        If Response = vbYes Then
       
            GPlayed = 0
            PlayerXTotal = 0
            PlayerOTotal = 0
            DrawTotal = 0
          End If
    Else
        
        MsgBox "There are no scores to erase!", vbCritical, "Error"
       
    End If
    'Displays the reseted scores
    lblGPlayed.Caption = Str$(GPlayed)
    lblPlayerX.Caption = Str$(PlayerXTotal)
    lblPlayerO.Caption = Str$(PlayerOTotal)
    lblDraws.Caption = Str$(DrawTotal)

End Sub
Private Sub lblTicTacToe_Click(Index As Integer)
    
    Dim GReset As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim DMSg As String
    
    'Checks if its an X or an O
    If TicNum Mod 2 = 1 And lblTicTacToe(Index).Caption <> "X" And lblTicTacToe(Index).Caption <> "O" Then
        lblTicTacToe(Index).Caption = "O"
        TicNum = TicNum + 1
        WinnerCount(Index) = 2
    ElseIf lblTicTacToe(Index).Caption <> "X" And lblTicTacToe(Index).Caption <> "O" Then
        lblTicTacToe(Index).Caption = "X"
        WinnerCount(Index) = 1
        TicNum = TicNum + 1
    End If
   
    'Determines the horizontal matches
    For X = 0 To 8 Step 3
        If WinnerCount(X) = WinnerCount(X + 1) And WinnerCount(X + 1) = WinnerCount(X + 2) Then
            Winner = WinnerCount(X)
        End If
    Next X
    
    'Determines vertical matches
    For Y = 0 To 2
        If WinnerCount(Y) = WinnerCount(Y + 3) And WinnerCount(Y + 3) = WinnerCount(Y + 6) Then
        
            Winner = WinnerCount(Y)
        
        End If
    Next Y
       
    'Determines the diagonal matches
    If WinnerCount(0) = WinnerCount(4) And WinnerCount(4) = WinnerCount(8) Then
        Winner = WinnerCount(0)
    ElseIf WinnerCount(2) = WinnerCount(4) And WinnerCount(4) = WinnerCount(6) Then
        Winner = WinnerCount(2)
    End If
    
    If Winner <> 0 Or TicNum = 9 Then
        'Determines the winner
        If Winner = 2 Then
            DMSg = "Player O wins!"
            PlayerOTotal = PlayerOTotal + 1
        ElseIf Winner = 1 Then
            DMSg = "Player X wins!"
            PlayerXTotal = PlayerXTotal + 1
    
        ElseIf TicNum = 9 Then
            DMSg = "No Winner!, draw"
            DrawTotal = DrawTotal + 1
        End If
        
        MsgBox DMSg, vbOKOnly + vbExclamation, "Results"
        'Resets the game
        For GReset = 0 To MAX
    
            lblTicTacToe(GReset).Caption = ""
            WinnerCount(GReset) = GReset + 3
    
        Next GReset
        GPlayed = GPlayed + 1
        Winner = 0
        TicNum = 0
        
        'Displays the scores
        lblGPlayed.Caption = Str$(GPlayed)
        lblPlayerX.Caption = Str$(PlayerXTotal)
        lblPlayerO.Caption = Str$(PlayerOTotal)
        lblDraws.Caption = Str$(DrawTotal)
    End If
     
End Sub
