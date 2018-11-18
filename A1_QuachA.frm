VERSION 5.00
Begin VB.Form frmA1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alex's sandy sandwich"
   ClientHeight    =   4530
   ClientLeft      =   3345
   ClientTop       =   2490
   ClientWidth     =   9120
   Icon            =   "A1_QuachA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   9120
   Begin VB.CheckBox chkDrink 
      Caption         =   "Soft Drink $0.99"
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
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CheckBox chkFrenchFries 
      Caption         =   "French Fries $1.59"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset the interface"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Place order"
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.Frame fraReceipt 
      Caption         =   "Your Reciept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   2535
      Begin VB.Label lblReceiptMsg 
         Caption         =   $"A1_QuachA.frx":030A
         Height          =   1215
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
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
         Left            =   1440
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblHST 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
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
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "$0.00"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "HST:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "SubTotal:"
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
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame fraMenu 
      Caption         =   "Order Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.CheckBox chkCombo 
         Caption         =   "Combo $5.29"
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
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
      End
      Begin VB.CheckBox chkSandwich 
         Caption         =   "Sandwich $3.69"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please note, if you would like all three (Sandwich, Fries and a drink), please select the combo option.  Thank you."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   8895
   End
   Begin VB.Label lblResturantName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome to Alex's sandy sandwich!  Please place your order below! "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8895
   End
End
Attribute VB_Name = "frmA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Alex Quach
'Date: 24/09/2015
'Purpose: To allow the user of the program to order
'what they want

Private Sub chkCombo_Click()

    'Checks to see if the user ordered a combo or not.
    If chkCombo.Value = 1 Then
        chkSandwich.Enabled = False
        chkFrenchFries.Enabled = False
        chkDrink.Enabled = False
        chkSandwich.Value = 1
        chkFrenchFries.Value = 1
        chkDrink.Value = 1
        chkCombo.Value = 1
    ElseIf chkCombo.Value = 0 Then
    
        chkSandwich.Enabled = True
        chkFrenchFries.Enabled = True
        chkDrink.Enabled = True
        chkSandwich.Value = 0
        chkFrenchFries.Value = 0
        chkDrink.Value = 0
    
    End If

End Sub

Private Sub chkDrink_Click()
    
    If chkSandwich.Value = 1 And chkDrink.Value = 1 And chkFrenchFries = 1 Then
    
        chkSandwich.Enabled = False
        chkFrenchFries.Enabled = False
        chkDrink.Enabled = False
        chkCombo.Value = 1
    
    End If

End Sub

Private Sub chkFrenchFries_Click()

    If chkSandwich.Value = 1 And chkDrink.Value = 1 And chkFrenchFries = 1 Then
    
        chkSandwich.Enabled = False
        chkFrenchFries.Enabled = False
        chkDrink.Enabled = False
        chkCombo.Value = 1
    
    End If

End Sub

Private Sub chkSandwich_Click()

    'Checks to see if the user clicks all three options.  If the user clicks all three options,
    'check off the combo option.
    If chkSandwich.Value = 1 And chkDrink.Value = 1 And chkFrenchFries = 1 Then
    
        chkSandwich.Enabled = False
        chkFrenchFries.Enabled = False
        chkDrink.Enabled = False
        chkCombo.Value = 1
    
    End If

End Sub

Private Sub cmdCalculate_Click()

    Const SANDWICH = 3.69
    Const FRIES = 1.59
    Const DRINK = 0.99
    Const HST = 0.13
    
    Dim SubTotal As Single
    Dim TaxTotal As Single
    Dim Total As Single
    
    SubTotal = 0
    TaxTotal = 0
    Total = 0
    
    'Checks to see if the user has not checked anything and clicked the
    'order button.
    If chkSandwich.Value = 0 And chkDrink.Value = 0 And chkFrenchFries = 0 Then
        
        lblSubTotal.Caption = Format$(0, "currency")
        lblHST.Caption = Format$(0, "currency")
        lblTotal.Caption = Format$(0, "currency")
        lblReceiptMsg.Visible = False
        MsgBox "You did not order anything! Please select a minimum of 1 item.", vbCritical, "Error!"
    
    Else
        'Checks to see the possible order combonations the user
        'may enter.
        If chkSandwich.Value = 1 Then
            
            SubTotal = SubTotal + SANDWICH
            
        End If
        
        If chkFrenchFries.Value = 1 Then
        
            SubTotal = SubTotal + FRIES
        
        End If
        
        If chkDrink.Value = 1 Then
        
            SubTotal = SubTotal + DRINK
        
        End If
        
        If chkCombo.Value = 1 Then
        
            SubTotal = 5.29
        
        End If
        
        'Checks to see if the HST (Tax) will be applied.
        
        If SubTotal >= 4 Then
        
            TaxTotal = SubTotal * HST
        
        End If
        
        'Final total calculation is made.
        Total = SubTotal + TaxTotal
        
        'Display the subtotal, tax total, the final total, and the reciept message.
        
        lblReceiptMsg.Visible = True
        lblSubTotal.Caption = Format$(SubTotal, "currency")
        lblHST.Caption = Format$(TaxTotal, "currency")
        lblTotal.Caption = Format$(Total, "currency")
    
    End If
    
    

End Sub

Private Sub cmdExit_Click()

    MsgBox "The program will now end.", vbInformation, "Program ending"
    End

End Sub

Private Sub cmdReset_Click()

    'Resets the entire interface.
    
    chkSandwich.Value = 0
    chkFrenchFries.Value = 0
    chkDrink.Value = 0
    chkCombo.Value = 0
    lblReceiptMsg.Visible = False
    lblSubTotal.Caption = Format$(0#, "currency")
    lblHST.Caption = Format$(0#, "currency")
    lblTotal.Caption = Format$(0#, "currency")

End Sub

Private Sub Form_Load()
    
    'Makes the receipt message that the user see when he/she orders not visible.
    lblReceiptMsg.Visible = False

End Sub
