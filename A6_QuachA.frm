VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Animation - Pacman"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   Icon            =   "A6_QuachA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgPacDown 
      Height          =   480
      Left            =   1920
      Picture         =   "A6_QuachA.frx":08CA
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacman 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image imgPacLeft 
      Height          =   480
      Left            =   720
      Picture         =   "A6_QuachA.frx":150C
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacRight 
      Height          =   480
      Left            =   120
      Picture         =   "A6_QuachA.frx":214E
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPacup 
      Height          =   480
      Left            =   1320
      Picture         =   "A6_QuachA.frx":2D90
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Alex Quach
'Date: February 23, 2016
'Purpose: To make a pacman move around the form
Option Explicit
Const MOVEMENTSPEED = 100
Dim CurrX As Integer
Dim CurrY As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

        'Checks to see what key is being pressed
        If KeyCode = vbKeyRight Then
            imgPacman.Picture = imgPacRight.Picture
            'Checks pacman location on form
            If imgPacman.Left <= frmMain.ScaleWidth Then
                CurrX = CurrX + MOVEMENTSPEED
                imgPacman.Move CurrX, CurrY
            Else
                'Wraps pacman
                CurrX = 0 - imgPacman.Width
                imgPacman.Move CurrX, CurrY
            End If
        ElseIf KeyCode = vbKeyLeft Then
            'Checks pacman location on form
            imgPacman.Picture = imgPacLeft.Picture
            If imgPacman.Left >= 0 - imgPacman.Width Then
                CurrX = CurrX - MOVEMENTSPEED
                imgPacman.Move CurrX, CurrY
            Else
                'Wraps pacman
                CurrX = frmMain.ScaleWidth
                imgPacman.Move CurrX, CurrY
            End If
        ElseIf KeyCode = vbKeyDown Then
            'Checks pacman location on form
            imgPacman.Picture = imgPacDown.Picture
            If imgPacman.Top <= frmMain.ScaleHeight Then
        
                CurrY = CurrY + MOVEMENTSPEED
                imgPacman.Move CurrX, CurrY
            Else
                'Wraps pacman
                CurrY = 0 - imgPacman.Height
                imgPacman.Move CurrX, CurrY
            End If
        ElseIf KeyCode = vbKeyUp Then
            'Checks pacman location on form
            imgPacman.Picture = imgPacup.Picture
            If imgPacman.Top >= 0 - imgPacman.Height Then
                CurrY = CurrY - MOVEMENTSPEED
                imgPacman.Move CurrX, CurrY
            Else
                'Wraps pacman
                CurrY = frmMain.ScaleHeight - imgPacman.Top
                imgPacman.Move CurrX, CurrY
        End If
    End If
    
End Sub

Private Sub Form_Load()
    'Intialize the pacman starting location
    CurrX = 0
    CurrY = 0
    imgPacman.Move CurrX, CurrY
    imgPacman.Picture = imgPacRight.Picture
End Sub
