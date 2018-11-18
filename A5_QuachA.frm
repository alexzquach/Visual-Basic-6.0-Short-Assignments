VERSION 5.00
Begin VB.Form frmAssignment4 
   Caption         =   "Riverdale Lottery 6/49"
   ClientHeight    =   3300
   ClientLeft      =   1605
   ClientTop       =   2025
   ClientWidth     =   7185
   Icon            =   "A5_QuachA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   7185
   Begin VB.CommandButton cmdWinningNumbers 
      Caption         =   "&Winning Numbers"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play ($2.00)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblBank 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Bank:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblMatch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label lblWinningNumbers 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1440
      Width           =   4935
   End
   Begin VB.Label Label4 
      Caption         =   "Numbers:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Winning"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblPlayerNumbers 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Numbers:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmAssignment4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MAXNUMBERS = 6
Const HIGH = 49
Const LOW = 1
Dim Bank As Long
Dim PlayerNumbers(1 To MAXNUMBERS) As Integer
Dim WinningNumbers(1 To MAXNUMBERS) As Integer
'Name: Alex Quach
'Date: January 22, 2016
'Purpose: To simulate a 6 number lottery game

Private Sub cmdExit_Click()
    
    'Exits the program
    Beep
    End
    
End Sub

Private Sub cmdPlay_Click()
    
    Dim PNumMsg As String
    Dim K As Integer
    Dim X As Integer
    Dim MsgCount As Integer
    Dim FNum As Integer
    Dim SNum As Integer
    Dim TempNum As Integer
    
    cmdWinningNumbers.Enabled = True
   
    
    'Checks if the player has enough money to play
    If Bank < 2 Then
        cmdPlay.Enabled = False
        MsgBox "Sorry, you do not have enough money to play!  Please end the program.", vbInformation, "No play"
        cmdWinningNumbers.Enabled = False
    Else
        'Clears the display
        PNumMsg = ""
        lblPlayerNumbers.Caption = ""
        lblWinningNumbers.Caption = ""
        lblMatch.Caption = ""
        'Subtracts $2 for each play if the player does have enough to play
        Bank = Bank - 2
        'Loops to randomize the player numbers
        For K = 1 To MAXNUMBERS
            PlayerNumbers(K) = Int(Rnd * (HIGH - LOW + 1) + LOW)
            For X = 1 To MAXNUMBERS
                If K <> X Then
                    Do While PlayerNumbers(K) = PlayerNumbers(X)
                        PlayerNumbers(K) = Int(Rnd * (HIGH - LOW + 1) + LOW)
                        If K <> 1 Then
                            X = 1
                        End If
                    Loop
                End If
            Next X
        Next K
        'Builds the player numbers message
        For MsgCount = 1 To MAXNUMBERS

            PNumMsg = PNumMsg & Str$(PlayerNumbers(MsgCount)) & " "
       
        Next MsgCount
        lblPlayerNumbers.Caption = PNumMsg
    
    End If
    'Displays the money the user currently has
    lblBank.Caption = Format$(Bank, "currency")
    cmdPlay.Enabled = False
  
    

End Sub

Private Sub cmdWinningNumbers_Click()
    
    Dim WNumMsg As String
    Dim CLoop As Integer
    Dim CLoopTwo As Integer
    Dim X As Integer
    Dim K As Integer
    Dim FNum As Integer
    Dim SNum As Integer
    Dim TempNum As Integer
    Dim Match As Integer
    
    WNumMsg = ""
    Match = 0

    'Loops to randomize the winning numbers
    For K = 1 To MAXNUMBERS
        WinningNumbers(K) = Int(Rnd * (HIGH - LOW + 1) + LOW)
        For X = 1 To MAXNUMBERS
            If K <> X Then
                Do While WinningNumbers(K) = WinningNumbers(X)
                    WinningNumbers(K) = Int(Rnd * (HIGH - LOW + 1) + LOW)
                    If K <> 1 Then
                        X = 1
                    End If
                Loop
            End If
        Next X
    Next K
     
    'Loops to compare the player numbers and winning numbers and builds the message
    For CLoop = 1 To MAXNUMBERS
        'Builds the winning number message
        WNumMsg = WNumMsg & Str$(WinningNumbers(CLoop)) & " "
        'Compares the player numbers and winning numbers to check if there are any matches
        For CLoopTwo = 1 To MAXNUMBERS
            If PlayerNumbers(CLoop) = WinningNumbers(CLoopTwo) Then
        
                Match = Match + 1
            
            End If
        Next CLoopTwo
        
    Next CLoop
    lblWinningNumbers.Caption = WNumMsg
    'Displays if the user wins or not and adds the appropriate prize money
    If Match > 2 Then
        lblMatch.Caption = "Congratualations, you win!" & Str$(Match) & " matches!"
        If Match = 3 Then
            Bank = Bank + 5
        ElseIf Match = 4 Then
            Bank = Bank + 100
        ElseIf Match = 5 Then
            Bank = Bank + 1000000
        ElseIf Match = 6 Then
            Bank = Bank + 20000000
        End If
    Else
    'Displays if the user loses
        lblMatch.Caption = "Better luck next time!" & Str$(Match) & " matches!"
    End If
    lblBank.Caption = Format$(Bank, "currency")
    cmdPlay.Enabled = True
    cmdWinningNumbers.Enabled = False
End Sub

Private Sub Form_Load()
    
    Dim K As Integer
    'Intialize the variables
    Randomize
    Bank = 20
    lblBank.Caption = Format$(Bank, "currency")
    cmdWinningNumbers.Enabled = False
    For K = 1 To MAXNUMBERS
    
        PlayerNumbers(K) = 0
        WinningNumbers(K) = 0
    
    Next K
   
End Sub
