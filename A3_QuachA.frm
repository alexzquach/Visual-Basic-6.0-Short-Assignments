VERSION 5.00
Begin VB.Form frmA3 
   Caption         =   "Wages "
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "A3_QuachA.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9720
      TabIndex        =   2
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdReadData 
      Caption         =   "&Read Data"
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   9435
      TabIndex        =   0
      Top             =   120
      Width           =   9495
   End
End
Attribute VB_Name = "frmA3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Name: Alex Quach
'Date: November 17, 2015
'Purpose: To read a data file and calculate the gross pay and the net pay for
'each employee in the data file

Private Sub cmdExit_Click()

    Beep
    End

End Sub

Private Sub cmdReadData_Click()
    Const DEDUCTION = 0.25
    Const OVERTIME = 1.5
    
    Dim EName As String
    Dim Wages As Single
    Dim Hours As Single
    Dim AverageNPay As Single
    Dim GPay As Single
    Dim NPay As Single
    Dim TotalNPay As Single
    Dim X As Integer
    
    picData.Cls
    
    TotalNPay = 0
    X = 0
    
    'Print sub headings.
    
    picData.Print Tab(5); "EMPLOYEE NAME"; Tab(35); "WAGE"; Tab(45);
    picData.Print "HOURS"; Tab(60); "GROSS PAY"; Tab(82); "NET PAY"
    picData.Print
    
    'Open the wages text file.
    
    Open App.Path & "\wages.txt" For Input As #1
    
    'Loop to read the data from the sequential file.
    Do While Not EOF(1)
    
        'Placing the data received into variables.
        Input #1, EName
        Input #1, Wages
        Input #1, Hours
        
        'Counting total number of employees.
        X = X + 1
        
        'Checking if the employee worked overtime or not and then
        'calculating the gross pay.
        If Hours <= 40 Then
            GPay = Wages * Hours
        Else
            GPay = (Wages * 40) + (Wages * (Hours - 40) * OVERTIME)
        End If
        
        'Calculating the net pay
        NPay = GPay - (GPay * DEDUCTION)
        TotalNPay = TotalNPay + NPay
        
        'Displaying the data to the user of the program.
        If X < 10 Then
            picData.Print Tab(2); X & "."; Tab(5); EName;
        ElseIf X >= 10 Then
            picData.Print Tab(1); X & "."; Tab(5); EName;
        End If
        picData.Print Tab(29); "$"; Tab(32);
        picData.Print Format$(Format$(Wages, "0.00"), "@@@@@@@");
        picData.Print Tab(45); Format$(Format$(Hours, "0.0"), "@@@@@");
        picData.Print Tab(55); "$";
        picData.Print Tab(58); Format$(Format$(GPay, "0.00"), "@@@@@@@@@@@");
        picData.Print Tab(75); "$"; Tab(78);
        picData.Print Format$(Format$(NPay, "0.00"), "@@@@@@@@@@@")
        
    Loop
    Close #1
    
    'Checking to see if there is data or not.
    If X <= 0 Then
        MsgBox "There is no information!", vbCritical, "Error!"
        MsgBox "The program will now exit.", vbCritical, "Exit"
        End
    Else
    'Displaying the final output.
        picData.Print
        picData.Print "Number of employees is: "; X
        AverageNPay = TotalNPay / X
        picData.Print "The average net pay is: "; Format$(AverageNPay, "currency")
    End If
End Sub
