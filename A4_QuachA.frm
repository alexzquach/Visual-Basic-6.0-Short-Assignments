VERSION 5.00
Begin VB.Form frmA4 
   Caption         =   "Player Information"
   ClientHeight    =   4020
   ClientLeft      =   4245
   ClientTop       =   2910
   ClientWidth     =   6465
   Icon            =   "A4_QuachA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6465
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblPNum 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Player:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblTotalPoints 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Total Points:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label lblAssists 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblGoals 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label lblGamesPlayed 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblTeamName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblFullName 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Assists:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Goals:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Games Played:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   " Team Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Full name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "frmA4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Quach
'Date: 9/12/2015
'Purpose: To read a data file and display the stats of
'the NHL player one at a time,
'using a next and previous button to switch between them

Option Explicit
Const MAXPLAYERS = 100
'Declare the variables whos value will be kept throughout
'the program
Dim X As Integer
Dim FullName(1 To MAXPLAYERS) As String
Dim TeamName(1 To MAXPLAYERS) As String
Dim GamesPlayed(1 To MAXPLAYERS) As Integer
Dim Goals(1 To MAXPLAYERS) As Integer
Dim Assists(1 To MAXPLAYERS) As Integer
Dim TotalPoints(1 To MAXPLAYERS) As Integer
Dim NumPlayers As Integer
Private Sub cmdExit_Click()
    'Exits the program
    Beep
    End

End Sub

Private Sub cmdNext_Click()
    'Increase the position of the array by one and display the data stored
    'in that position
    X = X + 1
    lblPNum.Caption = Str$(X) & " of " & Str$(NumPlayers)
    If Len(FullName(X)) > 16 Then
        lblFullName.Caption = Left$(FullName(X), 16) & "...."
    Else
        lblFullName.Caption = FullName(X)
    End If
    lblTeamName.Caption = TeamName(X)
    lblGamesPlayed.Caption = Str$(GamesPlayed(X))
    lblGoals.Caption = Str$(Goals(X))
    lblAssists.Caption = Str$(Assists(X))
    lblTotalPoints.Caption = Str$(TotalPoints(X))
    If X > 1 And X < NumPlayers Then
        cmdPrevious.Visible = True
    ElseIf X = NumPlayers Then
        cmdNext.Visible = False
    End If
    
End Sub

Private Sub cmdOpen_Click()

    Dim K As Integer
    X = 1
    K = 0
    'Opening the text file
    Open App.Path & "\NHLStats.txt" For Input As #1
    Do While Not EOF(1)
        'Read the file and putting the data into arrays
        K = K + 1
        Input #1, FullName(K)
        Input #1, TeamName(K)
        Input #1, GamesPlayed(K)
        Input #1, Goals(K)
        Input #1, Assists(K)
        'Caclulating the total amount of points
        TotalPoints(K) = Goals(K) + Assists(K)
    Loop
    Close #1
    NumPlayers = K
    'Display the data that is in the first index of the array
    If Len(FullName(X)) > 16 Then
        lblFullName.Caption = Left$(FullName(X), 16) & "...."
    Else
        lblFullName.Caption = FullName(X)
    End If
    lblPNum.Caption = Str$(X) & " of " & Str$(NumPlayers)
    lblTeamName.Caption = TeamName(X)
    lblGamesPlayed.Caption = Str$(GamesPlayed(X))
    lblGoals.Caption = Str$(Goals(X))
    lblAssists.Caption = Str$(Assists(X))
    lblTotalPoints.Caption = Str$(TotalPoints(X))
    
    cmdNext.Visible = True
    cmdPrevious.Visible = False
    

    
    
End Sub

Private Sub cmdPrevious_Click()
    'Decrease the position of the array by one and display
    'the values stored in that position
    X = X - 1
    If Len(FullName(X)) > 16 Then
        lblFullName.Caption = Left$(FullName(X), 16) & "...."
    Else
        lblFullName.Caption = FullName(X)
    End If
    
    lblPNum.Caption = Str$(X) & " of " & Str$(NumPlayers)
    lblTeamName.Caption = TeamName(X)
    lblGamesPlayed.Caption = Str$(GamesPlayed(X))
    lblGoals.Caption = Str$(Goals(X))
    lblAssists.Caption = Str$(Assists(X))
    lblTotalPoints.Caption = Str$(TotalPoints(X))
    If X > 1 And X < NumPlayers Then
        cmdPrevious.Visible = True
        cmdNext.Visible = True
    ElseIf X = 1 Then
        cmdPrevious.Visible = False
    End If

End Sub

Private Sub Form_Load()

    Dim K As Integer
    'Initialize the position of the array to one
    X = 1
    'Initialize the variables
    For K = 1 To MAXPLAYERS
        FullName(K) = ""
        TeamName(K) = ""
        GamesPlayed(K) = 0
        Goals(K) = 0
        Assists(K) = 0
        TotalPoints(K) = 0
    Next K
    
    NumPlayers = 0
    
    cmdNext.Visible = False
    cmdPrevious.Visible = False
    

End Sub
