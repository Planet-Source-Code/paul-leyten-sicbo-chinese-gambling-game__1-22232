VERSION 5.00
Begin VB.UserControl ucDices 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1860
   ScaleHeight     =   780
   ScaleWidth      =   1860
   Begin VB.Timer tmrRoll 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   480
      Top             =   720
   End
   Begin VB.Timer tmrStopRol 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   -30
      Top             =   720
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   0
      Top             =   0
      Width           =   1860
   End
   Begin VB.Image imgDice1 
      Height          =   495
      Index           =   1
      Left            =   120
      Picture         =   "ucDices.ctx":0000
      Top             =   105
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice1 
      Height          =   495
      Index           =   2
      Left            =   120
      Picture         =   "ucDices.ctx":0D26
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice1 
      Height          =   495
      Index           =   3
      Left            =   120
      Picture         =   "ucDices.ctx":2617
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice1 
      Height          =   495
      Index           =   4
      Left            =   120
      Picture         =   "ucDices.ctx":40EB
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice1 
      Height          =   495
      Index           =   5
      Left            =   120
      Picture         =   "ucDices.ctx":5D36
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice1 
      Height          =   495
      Index           =   6
      Left            =   120
      Picture         =   "ucDices.ctx":7AB3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice2 
      Height          =   495
      Index           =   1
      Left            =   720
      Picture         =   "ucDices.ctx":98C8
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice2 
      Height          =   495
      Index           =   2
      Left            =   720
      Picture         =   "ucDices.ctx":A5EE
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice2 
      Height          =   495
      Index           =   3
      Left            =   720
      Picture         =   "ucDices.ctx":BEDF
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice2 
      Height          =   495
      Index           =   4
      Left            =   720
      Picture         =   "ucDices.ctx":D9B3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice2 
      Height          =   495
      Index           =   5
      Left            =   720
      Picture         =   "ucDices.ctx":F5FE
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice2 
      Height          =   495
      Index           =   6
      Left            =   720
      Picture         =   "ucDices.ctx":1137B
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice3 
      Height          =   495
      Index           =   1
      Left            =   1320
      Picture         =   "ucDices.ctx":13190
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice3 
      Height          =   495
      Index           =   2
      Left            =   1320
      Picture         =   "ucDices.ctx":13EB6
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice3 
      Height          =   495
      Index           =   3
      Left            =   1320
      Picture         =   "ucDices.ctx":157A7
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice3 
      Height          =   495
      Index           =   4
      Left            =   1320
      Picture         =   "ucDices.ctx":1727B
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice3 
      Height          =   495
      Index           =   5
      Left            =   1320
      Picture         =   "ucDices.ctx":18EC6
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgDice3 
      Height          =   495
      Index           =   6
      Left            =   1320
      Picture         =   "ucDices.ctx":1AC43
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "ucDices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Dice(3) As Byte
Public Event Ready(Dice1 As Byte, Dice2 As Byte, Dice3 As Byte)
Public Event Click()

Public Sub RollDice()







    Call Inter_RollDice
    tmrRoll.Enabled = True
    tmrStopRol.Enabled = True

    Exit Sub







End Sub

Private Sub Image1_Click()







RaiseEvent Click

    Exit Sub







End Sub

Private Sub tmrRoll_Timer()







    Inter_RollDice

    Exit Sub







End Sub
Private Sub tmrStopRol_Timer()







    tmrRoll.Enabled = False
    tmrStopRol.Enabled = False
    RaiseEvent Ready(Dice(1), Dice(2), Dice(3))

    Exit Sub







End Sub
Private Sub Inter_RollDice()







    Randomize Timer
    If Dice(1) = 0 Or Dice(2) = 0 Or Dice(3) = 0 Then
    Else
        imgDice1(Dice(1)).Visible = False
        imgDice2(Dice(2)).Visible = False
        imgDice3(Dice(3)).Visible = False
    End If
    Dice(1) = Int((6 * Rnd) + 1)
    Dice(2) = Int((6 * Rnd) + 1)
    Dice(3) = Int((6 * Rnd) + 1)
    imgDice1(Dice(1)).Visible = True: imgDice1(Dice(1)).ZOrder 1
    imgDice2(Dice(2)).Visible = True: imgDice2(Dice(2)).ZOrder 1
    imgDice3(Dice(3)).Visible = True: imgDice3(Dice(3)).ZOrder 1

    Exit Sub







End Sub

Private Sub UserControl_Initialize()







    Dice(1) = Int((6 * Rnd) + 1)
    Dice(2) = Int((6 * Rnd) + 1)
    Dice(3) = Int((6 * Rnd) + 1)
    imgDice1(Dice(1)).Visible = True: imgDice1(Dice(1)).ZOrder 1
    imgDice2(Dice(2)).Visible = True: imgDice2(Dice(2)).ZOrder 1
    imgDice3(Dice(3)).Visible = True: imgDice3(Dice(3)).ZOrder 1

    Exit Sub







End Sub
