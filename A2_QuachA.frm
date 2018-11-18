VERSION 5.00
Begin VB.Form frmA2 
   Caption         =   "String Functions"
   ClientHeight    =   3420
   ClientLeft      =   4155
   ClientTop       =   4770
   ClientWidth     =   7785
   Icon            =   "A2_QuachA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   7785
   Begin VB.Frame Frame1 
      Caption         =   "Analysis Results:"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   7455
      Begin VB.Label lblNewPhrase 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label Label2 
         Caption         =   "Transmitted Phrase:"
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
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.TextBox txtPhrase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7335
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter a phrase:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmA2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author: Alex Quach
'Date: September 8th, 2015
'Purpose: The purpose of this program is to censor any four
'letter words the user entered.
Option Explicit


Private Sub cmdAnalyze_Click()

    Dim Phrase As String
    Dim NewWord As String
    Dim NewPhrase As String
    Dim Char As String
    Dim K As Integer

    NewPhrase = ""
    NewWord = ""
    Phrase = txtPhrase.Text
    
    For K = 1 To Len(Phrase)
    
        Char = Mid$(Phrase, K, 1)
        'Check if the character is a letter.
        If Char <> " " And Char <> "." Then
            NewWord = NewWord & Char
        Else
        'If the character is a space, check if it is a four letter word.
            If Len(NewWord) = 4 Then
                NewWord = "****"
            End If
        'Assign the word to the new phrase.
            NewPhrase = NewPhrase & NewWord & " "
            NewWord = ""
        
        End If
        
    Next K
    'Assign final word to the new phrase and displays the new phrase.
    NewPhrase = NewPhrase & NewWord
    lblNewPhrase.Caption = Left$(NewPhrase, Len(NewPhrase) - 1) & "."

End Sub
