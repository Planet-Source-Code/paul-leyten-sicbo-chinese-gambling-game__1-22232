VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSicBoPlayField 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SicBo"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "frmSicBoPlayfield.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSicBoPlayfield.frx":030A
   ScaleHeight     =   6390
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "That's IT"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   110
      Width           =   1000
   End
   Begin VB.Timer tmrPlayWin 
      Interval        =   100
      Left            =   4080
      Top             =   3000
   End
   Begin VB.PictureBox picDices 
      BackColor       =   &H008080FF&
      Height          =   645
      Left            =   3300
      ScaleHeight     =   585
      ScaleWidth      =   1725
      TabIndex        =   5
      Top             =   5595
      Visible         =   0   'False
      Width           =   1785
      Begin SicBo.ucDices ucDices 
         Height          =   1000
         Left            =   -100
         TabIndex        =   6
         Top             =   -60
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   1773
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000006&
      ForeColor       =   &H80000005&
      Height          =   645
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   5640
      Width           =   3060
   End
   Begin VB.Timer tmr_ShowWaitCursor 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   30
      Top             =   1395
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   5
      Left            =   2400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":4A38E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":4C06A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":4DD46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":4FA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":516FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":533DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   4
      Left            =   1800
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":550B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":56D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":58A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":5A74A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":5C426
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":5E102
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   3
      Left            =   1200
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":5FDDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":61ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":63796
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":65472
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":6714E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":68E2A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":6AB06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":6C7E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":6E4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":7019A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":71E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":73B52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   2
      Left            =   600
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":7582E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":7750A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":791E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":7AEC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":7CB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSicBoPlayfield.frx":7E87A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrShowPrice 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   -30
      Top             =   3420
   End
   Begin VB.Timer tmrRandomLights 
      Enabled         =   0   'False
      Interval        =   650
      Left            =   -30
      Top             =   3885
   End
   Begin VB.Timer tmr5Sec 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   3000
   End
   Begin VB.TextBox txtWin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3345
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   75
      Width           =   1275
   End
   Begin VB.TextBox txtCredits 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   75
      Width           =   1275
   End
   Begin VB.TextBox txtBet 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8040
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   75
      Width           =   1275
   End
   Begin VB.PictureBox picOneWinsFive 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      Picture         =   "frmSicBoPlayfield.frx":80556
      ScaleHeight     =   1500
      ScaleWidth      =   645
      TabIndex        =   3
      Top             =   3030
      Width           =   645
      Begin VB.Image Image2 
         Height          =   1515
         Left            =   0
         MousePointer    =   1  'Arrow
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Image Rules 
      Height          =   315
      Left            =   1215
      MousePointer    =   1  'Arrow
      Top             =   60
      Width           =   1200
   End
   Begin VB.Image Clear 
      Height          =   420
      Left            =   8400
      MousePointer    =   1  'Arrow
      Top             =   5760
      Width           =   855
   End
   Begin VB.Image picChip 
      Height          =   735
      Index           =   5
      Left            =   7770
      Top             =   5565
      Width           =   660
   End
   Begin VB.Image picChip 
      Height          =   720
      Index           =   4
      Left            =   7095
      Top             =   5580
      Width           =   675
   End
   Begin VB.Image picChip 
      Height          =   690
      Index           =   3
      Left            =   6420
      Top             =   5580
      Width           =   645
   End
   Begin VB.Image picChip 
      Height          =   705
      Index           =   2
      Left            =   5820
      Top             =   5580
      Width           =   600
   End
   Begin VB.Image picChip 
      Height          =   735
      Index           =   1
      Left            =   5115
      Top             =   5565
      Width           =   705
   End
   Begin VB.Image RollTheDice 
      Height          =   735
      Left            =   3240
      MousePointer    =   1  'Arrow
      Top             =   5565
      Width           =   1845
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   30
      MousePointer    =   1  'Arrow
      Top             =   5190
      Width           =   9360
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   1365
      MousePointer    =   1  'Arrow
      Top             =   390
      Width           =   6675
   End
   Begin VB.Image picSingle 
      Height          =   660
      Index           =   4
      Left            =   4695
      Tag             =   "Four"
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Image picAnyTriple 
      Height          =   1305
      Left            =   4410
      Tag             =   "Any triple "
      Top             =   750
      Width           =   570
   End
   Begin VB.Image picDouble 
      Height          =   1305
      Index           =   1
      Left            =   1365
      Tag             =   "Double One"
      Top             =   750
      Width           =   555
   End
   Begin VB.Image picDouble 
      Height          =   1305
      Index           =   2
      Left            =   1920
      Tag             =   "Double Two"
      Top             =   750
      Width           =   600
   End
   Begin VB.Image picDouble 
      Height          =   1305
      Index           =   3
      Left            =   2520
      Tag             =   "Double Three"
      Top             =   750
      Width           =   570
   End
   Begin VB.Image picTriple 
      Height          =   435
      Index           =   1
      Left            =   3090
      Tag             =   "Triple One"
      Top             =   750
      Width           =   1320
   End
   Begin VB.Image picTriple 
      Height          =   435
      Index           =   2
      Left            =   3090
      Tag             =   "Triple Two"
      Top             =   1185
      Width           =   1320
   End
   Begin VB.Image picTriple 
      Height          =   435
      Index           =   3
      Left            =   3090
      Tag             =   "Triple Three"
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image picTriple 
      Height          =   435
      Index           =   4
      Left            =   4980
      Tag             =   "Triple Four"
      Top             =   750
      Width           =   1320
   End
   Begin VB.Image picTriple 
      Height          =   435
      Index           =   5
      Left            =   4980
      Tag             =   "Triple Five"
      Top             =   1185
      Width           =   1320
   End
   Begin VB.Image picTriple 
      Height          =   435
      Index           =   6
      Left            =   4980
      Tag             =   "Triple Six"
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image picDouble 
      Height          =   1305
      Index           =   4
      Left            =   6300
      Tag             =   "Double Four"
      Top             =   750
      Width           =   570
   End
   Begin VB.Image picDouble 
      Height          =   1305
      Index           =   5
      Left            =   6870
      Tag             =   "Double Five"
      Top             =   750
      Width           =   585
   End
   Begin VB.Image picDouble 
      Height          =   1305
      Index           =   6
      Left            =   7455
      Tag             =   "Double Six"
      Top             =   750
      Width           =   585
   End
   Begin VB.Image picSmall 
      Height          =   1605
      Left            =   40
      Tag             =   "Small"
      Top             =   450
      Width           =   1320
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   4
      Left            =   0
      Tag             =   "Sum of eyes is 4"
      Top             =   2055
      Width           =   705
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   5
      Left            =   705
      Tag             =   "Sum of eyes is 5"
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   6
      Left            =   1380
      Tag             =   "Sum of eyes is 6"
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   7
      Left            =   2055
      Tag             =   "Sum of eyes is 7"
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   8
      Left            =   2715
      Tag             =   "Sum of eyes is 8"
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   9
      Left            =   3375
      Tag             =   "Sum of eyes is 9"
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   10
      Left            =   4050
      Tag             =   "Sum of eyes is 10"
      Top             =   2055
      Width           =   645
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   11
      Left            =   4695
      Tag             =   "Sum of eyes is 11"
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   12
      Left            =   5370
      Tag             =   "Sum of eyes is 12"
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   17
      Left            =   8685
      Tag             =   "Sum of eyes is 17"
      Top             =   2055
      Width           =   750
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   16
      Left            =   8025
      Tag             =   "Sum of eyes is 16"
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   15
      Left            =   7365
      Tag             =   "Sum of eyes is 15"
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   14
      Left            =   6690
      Tag             =   "Sum of eyes is 14"
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image picCount 
      Height          =   975
      Index           =   13
      Left            =   6030
      Tag             =   "Sum of eyes is 13"
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   12
      Left            =   645
      Tag             =   "One and Two"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   13
      Left            =   1215
      Tag             =   "One and Three"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   14
      Left            =   1800
      Tag             =   "One and Four"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   15
      Left            =   2370
      Tag             =   "One and Five"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   16
      Left            =   2940
      Tag             =   "One and Six"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   23
      Left            =   3510
      Tag             =   "Two and Three"
      Top             =   3030
      Width           =   600
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   24
      Left            =   4110
      Tag             =   "Two and Four"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   26
      Left            =   5265
      Tag             =   "Two and Six"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   34
      Left            =   5835
      Tag             =   "Three and Four"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   35
      Left            =   6420
      Tag             =   "Three and Five"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   36
      Left            =   7005
      Tag             =   "Three and Six"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   45
      Left            =   7575
      Tag             =   "Four and Five"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   46
      Left            =   8160
      Tag             =   "Four and Six"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   56
      Left            =   8745
      Tag             =   "Five and Six"
      Top             =   3030
      Width           =   675
   End
   Begin VB.Image picSingle 
      Height          =   660
      Index           =   1
      Left            =   0
      Tag             =   "One"
      Top             =   4530
      Width           =   1605
   End
   Begin VB.Image picSingle 
      Height          =   660
      Index           =   2
      Left            =   1605
      Tag             =   "Two"
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Image picSingle 
      Height          =   660
      Index           =   3
      Left            =   3150
      Tag             =   "Three"
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Image picSingle 
      Height          =   660
      Index           =   5
      Left            =   6240
      Tag             =   "Five"
      Top             =   4530
      Width           =   1560
   End
   Begin VB.Image picSingle 
      Height          =   660
      Index           =   6
      Left            =   7785
      Tag             =   "Six"
      Top             =   4530
      Width           =   1635
   End
   Begin VB.Image picBig 
      Height          =   1590
      Left            =   8040
      Tag             =   "Big"
      Top             =   465
      Width           =   1305
   End
   Begin VB.Image picCombi 
      Height          =   1500
      Index           =   25
      Left            =   4680
      Tag             =   "Two and Five"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   2415
      MousePointer    =   1  'Arrow
      Top             =   30
      Width           =   9270
   End
   Begin VB.Image imgCombiFiveAndSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFiveAndSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFiveAndSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFiveAndSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFiveAndSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourAndFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiThreeAndFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndThreeBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndThreeBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndThreeBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndThreeBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiTwoAndThreeBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourteenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourteenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourteenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourteenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiFourteenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndThreeBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndThreeBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndThreeBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndThreeBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndThreeBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndTwoBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndTwoBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndTwoBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndTwoBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCombiOneAndTwoBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSeventeenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSeventeenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSeventeenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSeventeenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSeventeenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixteenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixteenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixteenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixteenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixteenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFifthteenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFifthteenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFifthteenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFifthteenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFifthteenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourteenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourteenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourteenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourteenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourteenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountThirteenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountThirteenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountThirteenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountThirteenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountThirteenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTwelveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTwelveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTwelveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTwelveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTwelveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountElevenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountElevenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountElevenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountElevenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountElevenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountTenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountNineBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountNineBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountNineBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountNineBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountNineBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountEightBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountEightBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountEightBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountEightBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountEightBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSevenBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSevenBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSevenBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSevenBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSevenBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgCountFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleThreeBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleThreeBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleThreeBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleThreeBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleThreeBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleTwoBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleTwoBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleTwoBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleTwoBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleTwoBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleOneBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleOneBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleOneBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleOneBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSingleOneBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleThreeBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleThreeBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleThreeBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleThreeBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleThreeBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleTwoBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleTwoBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleTwoBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleTwoBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleTwoBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleOneBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleOneBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleOneBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleOneBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleOneBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSmallBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSmallBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSmallBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSmallBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSmallBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgAnyTripleBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgAnyTripleBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgAnyTripleBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgAnyTripleBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgAnyTripleBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleOneBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleOneBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleOneBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleOneBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleOneBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleTwoBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleTwoBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleTwoBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleTwoBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleTwoBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleThreeBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleThreeBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleThreeBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleThreeBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleThreeBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleAnyBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleAnyBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleAnyBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleAnyBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleAnyBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgTripleSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFourBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFourBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFourBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFourBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFourBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFiveBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFiveBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFiveBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFiveBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleFiveBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleSixBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleSixBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleSixBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleSixBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgDoubleSixBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgBigBet 
      Height          =   585
      Index           =   1
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgBigBet 
      Height          =   585
      Index           =   2
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgBigBet 
      Height          =   585
      Index           =   3
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgBigBet 
      Height          =   585
      Index           =   4
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgBigBet 
      Height          =   585
      Index           =   5
      Left            =   -30
      Top             =   585
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image imgSmall 
      Height          =   1605
      Left            =   40
      Picture         =   "frmSicBoPlayfield.frx":8243D
      Top             =   450
      Width           =   1320
   End
   Begin VB.Image imgDouble 
      Height          =   1305
      Index           =   1
      Left            =   1365
      Picture         =   "frmSicBoPlayfield.frx":8537D
      Top             =   750
      Width           =   555
   End
   Begin VB.Image imgDouble 
      Height          =   1305
      Index           =   2
      Left            =   1920
      Picture         =   "frmSicBoPlayfield.frx":86EF9
      Top             =   750
      Width           =   600
   End
   Begin VB.Image imgDouble 
      Height          =   1305
      Index           =   3
      Left            =   2520
      Picture         =   "frmSicBoPlayfield.frx":88B58
      Top             =   750
      Width           =   570
   End
   Begin VB.Image imgTriple 
      Height          =   435
      Index           =   1
      Left            =   3090
      Picture         =   "frmSicBoPlayfield.frx":8A857
      Top             =   750
      Width           =   1320
   End
   Begin VB.Image imgTriple 
      Height          =   435
      Index           =   2
      Left            =   3090
      Picture         =   "frmSicBoPlayfield.frx":8C3EA
      Top             =   1185
      Width           =   1320
   End
   Begin VB.Image imgTriple 
      Height          =   435
      Index           =   3
      Left            =   3090
      Picture         =   "frmSicBoPlayfield.frx":8E039
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image imgAnyTriple 
      Height          =   1305
      Left            =   4410
      Picture         =   "frmSicBoPlayfield.frx":8FCF0
      Top             =   750
      Width           =   570
   End
   Begin VB.Image imgTriple 
      Height          =   435
      Index           =   4
      Left            =   4980
      Picture         =   "frmSicBoPlayfield.frx":91E48
      Top             =   750
      Width           =   1320
   End
   Begin VB.Image imgTriple 
      Height          =   435
      Index           =   5
      Left            =   4980
      Picture         =   "frmSicBoPlayfield.frx":93AF3
      Top             =   1185
      Width           =   1320
   End
   Begin VB.Image imgTriple 
      Height          =   435
      Index           =   6
      Left            =   4980
      Picture         =   "frmSicBoPlayfield.frx":957FF
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Image imgDouble 
      Height          =   1305
      Index           =   4
      Left            =   6300
      Picture         =   "frmSicBoPlayfield.frx":97487
      Top             =   750
      Width           =   570
   End
   Begin VB.Image imgDouble 
      Height          =   1305
      Index           =   5
      Left            =   6870
      Picture         =   "frmSicBoPlayfield.frx":991EB
      Top             =   750
      Width           =   585
   End
   Begin VB.Image imgDouble 
      Height          =   1305
      Index           =   6
      Left            =   7455
      Picture         =   "frmSicBoPlayfield.frx":9AF9F
      Top             =   750
      Width           =   585
   End
   Begin VB.Image imgBig 
      Height          =   1590
      Left            =   8040
      Picture         =   "frmSicBoPlayfield.frx":9CD2A
      Top             =   465
      Width           =   1305
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   4
      Left            =   0
      Picture         =   "frmSicBoPlayfield.frx":9F9EA
      Top             =   2055
      Width           =   705
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   5
      Left            =   705
      Picture         =   "frmSicBoPlayfield.frx":A1425
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   6
      Left            =   1380
      Picture         =   "frmSicBoPlayfield.frx":A2DBC
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   7
      Left            =   2055
      Picture         =   "frmSicBoPlayfield.frx":A4750
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   8
      Left            =   2715
      Picture         =   "frmSicBoPlayfield.frx":A5FFF
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   9
      Left            =   3375
      Picture         =   "frmSicBoPlayfield.frx":A79D2
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   10
      Left            =   4050
      Picture         =   "frmSicBoPlayfield.frx":A946E
      Top             =   2055
      Width           =   645
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   11
      Left            =   4695
      Picture         =   "frmSicBoPlayfield.frx":AAF1A
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   12
      Left            =   5370
      Picture         =   "frmSicBoPlayfield.frx":AC8E2
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   13
      Left            =   6030
      Picture         =   "frmSicBoPlayfield.frx":AE372
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   14
      Left            =   6690
      Picture         =   "frmSicBoPlayfield.frx":AFD82
      Top             =   2055
      Width           =   675
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   15
      Left            =   7365
      Picture         =   "frmSicBoPlayfield.frx":B1752
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   16
      Left            =   8025
      Picture         =   "frmSicBoPlayfield.frx":B3016
      Top             =   2055
      Width           =   660
   End
   Begin VB.Image imgCount 
      Height          =   975
      Index           =   17
      Left            =   8685
      Picture         =   "frmSicBoPlayfield.frx":B494E
      Top             =   2055
      Width           =   750
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   56
      Left            =   8745
      Picture         =   "frmSicBoPlayfield.frx":B62F2
      Tag             =   "Even"
      Top             =   3030
      Width           =   675
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   46
      Left            =   8160
      Picture         =   "frmSicBoPlayfield.frx":B812E
      Tag             =   "Odd"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   45
      Left            =   7575
      Picture         =   "frmSicBoPlayfield.frx":B9F2A
      Tag             =   "Even"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   36
      Left            =   7005
      Picture         =   "frmSicBoPlayfield.frx":BBD15
      Tag             =   "Odd"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   35
      Left            =   6420
      Picture         =   "frmSicBoPlayfield.frx":BDAE5
      Tag             =   "Even"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   34
      Left            =   5835
      Picture         =   "frmSicBoPlayfield.frx":BF8A4
      Tag             =   "Odd"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   26
      Left            =   5265
      Picture         =   "frmSicBoPlayfield.frx":C16A0
      Tag             =   "Even"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   25
      Left            =   4680
      Picture         =   "frmSicBoPlayfield.frx":C345C
      Tag             =   "Odd"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   24
      Left            =   4110
      Picture         =   "frmSicBoPlayfield.frx":C51E3
      Tag             =   "Even"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   23
      Left            =   3510
      Picture         =   "frmSicBoPlayfield.frx":C6F73
      Tag             =   "Odd"
      Top             =   3030
      Width           =   600
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   16
      Left            =   2940
      Picture         =   "frmSicBoPlayfield.frx":C8C7B
      Tag             =   "Even"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   15
      Left            =   2370
      Picture         =   "frmSicBoPlayfield.frx":CAA36
      Tag             =   "Odd"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   14
      Left            =   1800
      Picture         =   "frmSicBoPlayfield.frx":CC7AE
      Tag             =   "Even"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   13
      Left            =   1215
      Picture         =   "frmSicBoPlayfield.frx":CE475
      Tag             =   "Odd"
      Top             =   3030
      Width           =   585
   End
   Begin VB.Image imgCombi 
      Height          =   1500
      Index           =   12
      Left            =   645
      Picture         =   "frmSicBoPlayfield.frx":D00BC
      Tag             =   "Even"
      Top             =   3030
      Width           =   570
   End
   Begin VB.Image imgSingle 
      Height          =   660
      Index           =   1
      Left            =   0
      Picture         =   "frmSicBoPlayfield.frx":D1CDF
      Top             =   4530
      Width           =   1605
   End
   Begin VB.Image imgSingle 
      Height          =   660
      Index           =   2
      Left            =   1605
      Picture         =   "frmSicBoPlayfield.frx":D388F
      Tag             =   "Two"
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Image imgSingle 
      Height          =   660
      Index           =   3
      Left            =   3150
      Picture         =   "frmSicBoPlayfield.frx":D538B
      Tag             =   "Three"
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Image imgSingle 
      Height          =   660
      Index           =   4
      Left            =   4695
      Picture         =   "frmSicBoPlayfield.frx":D707B
      Tag             =   "Four"
      Top             =   4530
      Width           =   1545
   End
   Begin VB.Image imgSingle 
      Height          =   660
      Index           =   5
      Left            =   6240
      Picture         =   "frmSicBoPlayfield.frx":D8D43
      Tag             =   "Five"
      Top             =   4530
      Width           =   1560
   End
   Begin VB.Image imgSingle 
      Height          =   660
      Index           =   6
      Left            =   7785
      Picture         =   "frmSicBoPlayfield.frx":DA957
      Tag             =   "Six"
      Top             =   4530
      Width           =   1635
   End
End
Attribute VB_Name = "frmSicBoPlayField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum sbLights
    sbRandom = 0
    sbOff = 1
    sbOn = 2
End Enum
Dim rndLights As Boolean
Dim DragChip As Byte
Dim Game As New clsGame
Private Dice(3) As Byte
Dim ds As DirectSound
Dim dsb As DirectSoundBuffer
Sub MoveChip(Button As Integer, xLeft As Single, yTop As Single, imgSource As Image, imgChip As Image, Tag As String)
Reset
CreateDSBFromWaveFile ds, App.Path & "\click.wav", dsb
dsb.Play 0, 0, lngFlag
Dim sKey As String
    Dim bet As Double
    Dim NewKey As Boolean
    Dim MultipleChips As Boolean
yTop = imgSource.Top + yTop - 280 '160
xLeft = imgSource.Left + xLeft - 280 '- 160
If Button = 2 Then
    Me.MousePointer = vbNormal
    Exit Sub
End If
If Game.CreditsLeft - Key2Value(DragChip) < 0 Then
    Me.MousePointer = vbNormal
    Exit Sub
End If
sKey = imgSource.Name & GetIndex(imgSource)
If Game.colBets.Count = 0 Then ' altijd nieuw!
    Game.colBets.Add Key2Value(DragChip), xLeft, yTop, imgChip, sKey
Else ' collectie is niet leeg.
    NewKey = True
    For iCount = 1 To Game.colBets.Count
        If Game.colBets.item(iCount).sKey = sKey Then NewKey = False
    Next
    If NewKey Then
        Game.colBets.Add Key2Value(DragChip), xLeft, yTop, imgChip, sKey
    Else
        'Updaten van bet....
        'zoek ook uit of de juiste chip er al instaat (imgcol!)
        NewKey = True
        With Game.colBets.item(sKey).colImg
            For iCount = 1 To .Count
                If .item(iCount).sKey = bet2sKey(Key2Value(DragChip)) Then NewKey = False
            Next
            If NewKey Then
                .Add xLeft, yTop, imgChip, bet2sKey(Key2Value(DragChip))
            Else
                Game.colBets.item(sKey).colImg.item(bet2sKey(Key2Value(DragChip))).AddChip
            End If
        End With
        Game.colBets.item(sKey).bet = Game.colBets.item(sKey).bet + Key2Value(DragChip)
    End If
End If
iCount = Game.colBets.item(sKey).colImg.item(bet2sKey(Key2Value(DragChip))).CountChips
If iCount > 6 Then
    If iCount Mod 2 = 0 Then
        iCount = 6
    Else
        iCount = 5
    End If
End If
Set imgChip.Picture = ImageList1(DragChip).ListImages(iCount).Picture
    imgChip.Left = Game.colBets.item(sKey).colImg.item(bet2sKey(Key2Value(DragChip))).imgLeft
    imgChip.Top = Game.colBets.item(sKey).colImg.item(bet2sKey(Key2Value(DragChip))).imgTop
    imgChip.Visible = True
    
    bet = Key2Value(DragChip)
    Game.Add2Bet bet
    
    List1.AddItem "Your bet is: " & Key2Value(DragChip) & " at: " & Tag
    List1.ListIndex = List1.NewIndex
    
    Exit Sub
End Sub
Private Sub Clear_Click()
List1.Clear
Dim iCreditsLeft As Double
Dim blnStop As Boolean
Dim iCount As Integer
    'give the bets back!
    If Game.CurrentBet > 0 Then
        Game.CreditsLeft = Game.CreditsLeft + Game.CurrentBet
        iCreditsLeft = Game.CreditsLeft
        'remove pictures
        iCount = Game.colBets.Count
        If iCount > 0 Then
            While Not blnStop
                Game.colBets.Remove iCount
                iCount = Game.colBets.Count
                If iCount <= 0 Then blnStop = True
            Wend
        End If
    
        Set Game = Nothing
        Set Game = New clsGame
        Set Game.Form = Me
        Game.CreditsLeft = iCreditsLeft
        Game.StartGame
    End If
    Reset
    Exit Sub
End Sub
Sub Lights(Status As sbLights)
    Dim blnVisible As Boolean
    Select Case Status
        Case sbRandom
            tmr5Sec.Enabled = True
            tmrRandomLights.Enabled = True
            Exit Sub
        Case sbOn
            blnVisible = True
        Case sbOff
            blnVisible = False
    End Select
    imgBig.Visible = blnVisible
    imgSmall.Visible = blnVisible
    imgAnyTriple.Visible = blnVisible
    For Each item In imgDouble
        imgDouble(item.Index).Visible = blnVisible
    Next
    For Each item In imgTriple
        imgTriple(item.Index).Visible = blnVisible
    Next
    For Each item In imgCombi
        imgCombi(item.Index).Visible = blnVisible
    Next
    For Each item In imgCount
        imgCount(item.Index).Visible = blnVisible
    Next
    For Each item In imgSingle
        imgSingle(item.Index).Visible = blnVisible
    Next
    Exit Sub
End Sub
Sub Random(blnVisible As Boolean)
    imgBig.Visible = blnVisible
    imgSmall.Visible = blnVisible
    imgAnyTriple.Visible = Not blnVisible
    For Each item In imgDouble
        imgDouble(item.Index).Visible = Not blnVisible
    Next
    For Each item In imgTriple
        imgTriple(item.Index).Visible = blnVisible
    Next
    For Each item In imgCombi
        If item.Tag = "Even" Then
            imgCombi(item.Index).Visible = blnVisible
        Else
            imgCombi(item.Index).Visible = Not blnVisible
        End If
    Next
    For Each item In imgCount
        If item.Index Mod 2 = 0 Then
            imgCount(item.Index).Visible = blnVisible
        Else
            imgCount(item.Index).Visible = Not blnVisible
        End If
    Next
    For Each item In imgSingle
        If item.Index Mod 2 = 0 Then
            imgSingle(item.Index).Visible = blnVisible
        Else
            imgSingle(item.Index).Visible = Not blnVisible
        End If
    Next
    Exit Sub
End Sub

Private Sub cmdEnd_Click()
  Unload Me
End Sub

Private Sub Form_Click()
    Reset
    Exit Sub
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 And DragChip Then
    Me.MousePointer = vbNormal
End If
    Exit Sub
End Sub
Private Sub Form_Load()
  On Error GoTo Err_Sound
    DirectSoundCreate ByVal 0&, ds, Nothing
    ds.SetCooperativeLevel Me.hwnd, DSSCL_NORMAL
    
  On Error GoTo Hell
    Set Game = New clsGame
    Set Game.Form = Me
    Game.CreditsLeft = 500
    Game.StartGame
    tmrRandomLights.Enabled = True
    Exit Sub
Err_Sound:
  MsgBox "Direct Sound Object could not be created... sorry" & vbCrLf & "If you want to try this code without sound... you have to program it :-))"
  Unload Me

Hell:
  ' very shitty err_handling....
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not dsb Is Nothing Then
    Set dsb = Nothing
    Set ds = Nothing
End If
    Exit Sub
End Sub
Private Sub Image1_Click()
If MsgBox("Would you really like to Start over ?", vbYesNo) = vbYes Then
    Set Game = Nothing
    Set Game = New clsGame
    Set Game.Form = Me
    Game.CreditsLeft = 500
    Game.StartGame
End If
    Exit Sub
End Sub
Private Sub imgAnyTriple_Click()
    Reset
    Exit Sub
End Sub
Private Sub imgBig_Click()
    Reset
    Exit Sub
End Sub
Private Sub imgCombi_Click(Index As Integer)
    Reset
    Exit Sub
End Sub
Private Sub imgCount_Click(Index As Integer)
    Reset
    Exit Sub
End Sub
Private Sub imgDouble_Click(Index As Integer)
    Reset
    Exit Sub
End Sub
Private Sub imgSingle_Click(Index As Integer)
    Reset
    Exit Sub
End Sub
Private Sub imgSmall_Click()
    Reset
    Exit Sub
End Sub
Private Sub imgTriple_Click(Index As Integer)
    Reset
    Exit Sub
End Sub
Private Sub Lobby_Click()

End Sub
Private Sub RollTheDice_Click()
Screen.MousePointer = vbHourglass
    Reset
    If Game.colBets.Count > 0 Then
        Call Lights(sbOff)
        Dim lngFlag As Long
        If Not dsb Is Nothing Then
            dsb.Stop
            Set dsb = Nothing
        End If
        RollTheDice.Visible = True
        CreateDSBFromWaveFile ds, App.Path & "\dice.wav", dsb
        dsb.Play 0, 0, lngFlag
        picDices.Visible = True
        
        ucDices.RollDice
    End If
    Exit Sub
End Sub
Private Sub Rules_Click()
  Screen.MousePointer = vbHourglass
  tmr_ShowWaitCursor.Enabled = True
  Dim B As Long
  B = ShellExecute(0, "open", App.Path & "\sicbo_help.html", "", "", 1)
  Exit Sub
End Sub
Private Sub Timer1_Timer()
    Exit Sub
End Sub
Private Sub tmr_ShowWaitCursor_Timer()
Screen.MousePointer = vbNormal
tmr_ShowWaitCursor.Enabled = False
    Exit Sub
End Sub
Private Sub tmr5Sec_Timer()
    tmrRandomLights.Enabled = False
    tmr5Sec.Enabled = False
    Exit Sub
End Sub
Private Sub tmrPlayWin_Timer()
        DoEvents
        CreateDSBFromWaveFile ds, App.Path & "\Win.wav", dsb
        dsb.Play 0, 0, lngFlag
        tmrPlayWin.Enabled = False
    Exit Sub
End Sub
Private Sub tmrRandomLights_Timer()
    picDices.Visible = False
    Call Random(rndLights)
    rndLights = Not rndLights
    Exit Sub
End Sub
Private Sub tmrShowPrice_Timer()
    tmrRandomLights.Enabled = True
    txtWin = ""
    Exit Sub
End Sub
Sub Reset()
    picDices.Visible = False
    txtWin = ""
    txtCredits = Game.CreditsLeft
    tmrRandomLights.Enabled = False
    tmrShowPrice.Enabled = False
    Call Lights(sbOn)
    Exit Sub
End Sub
Private Sub txtBet_Change()
    Exit Sub
End Sub
Private Sub txtCredits_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = vbShiftMask And KeyCode = 13 Then
  Game.CreditsLeft = txtCredits
End If
    Exit Sub
End Sub
Private Sub ucDices_Click()
    picDices.Visible = False
    Exit Sub
End Sub
Private Sub ucDices_Ready(Dice1 As Byte, Dice2 As Byte, Dice3 As Byte)
    Dim iWin As Double
    Dim strWin  As String
    iWin = Game.CheckWins(Dice1, Dice2, Dice3)
'    iWin = Game.CheckWins(3, 3, 4)
    Game.CreditsLeft = Game.CreditsLeft + iWin
    tmrShowPrice.Enabled = True
    If iWin > 0 Then
        strWin = "You won: " & iWin
    Else
        strWin = "You didn't win...."
    End If
    List1.AddItem "Winning Dices: " & Dice1 & "," & Dice2 & "," & Dice3 & ", " & strWin
    List1.ListIndex = List1.NewIndex
    Me.SetFocus
    dsb.Stop
    Set dsb = Nothing
    If iWin > 0 Then tmrPlayWin.Enabled = True
    Screen.MousePointer = vbNormal
    Exit Sub
End Sub
Public Sub CreateDSBFromWaveFile(ds As DirectSound, ByVal strFile As String, dsb As DirectSoundBuffer)
    Dim hWave As Long
    Dim pcmwave As WAVEFORMATEX
    Dim lngSize As Long
    Dim lngPosition As Long
    Dim ptr1 As Long, ptr2 As Long, lng1 As Long, lng2 As Long
    Dim aByte() As Byte
    
    ' Byte array to load the whole file
    ReDim aByte(1 To FileLen(strFile))
    hWave = FreeFile
    Open strFile For Binary As hWave
    
    ' Load the whole file in the byte array
    Get hWave, , aByte
    Close hWave
    
    ' Search "fmt" tag
    lngPosition = 1
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) <> "fmt"
        lngPosition = lngPosition + 1
    Wend
    
    ' Copy wave header to structure
    CopyMemory VarPtr(pcmwave), VarPtr(aByte(lngPosition + 8)), Len(pcmwave)
    
    ' Search "data" tag
    While Chr$(aByte(lngPosition)) + Chr$(aByte(lngPosition + 1)) + Chr$(aByte(lngPosition + 2)) + Chr$(aByte(lngPosition + 3)) <> "data"
        lngPosition = lngPosition + 1
    Wend
    
    ' Get the data size
    CopyMemory VarPtr(lngSize), VarPtr(aByte(lngPosition + 4)), Len(lngSize)
    
    ' Fill buffer description
    Dim dsbd As DSBUFFERDESC
    With dsbd
        .dwSize = Len(dsbd)
        .dwFlags = DSBCAPS_CTRLDEFAULT 'Or DSBCAPS_STATIC Or DSBCAPS_LOCSOFTWARE
        .dwBufferBytes = lngSize
        .lpwfxFormat = VarPtr(pcmwave)
    End With
    
    ' Create the sound buffer
    ds.CreateSoundBuffer dsbd, dsb, Nothing
    
    ' Lock
    dsb.Lock 0&, lngSize, ptr1, lng1, ptr2, lng2, 0&
    
    ' Copy data to buffer
    CopyMemory ptr1, VarPtr(aByte(lngPosition + 4 + 4)), lng1
    
    ' Copy second part if needed
    If lng2 <> 0 Then
        CopyMemory ptr2, VarPtr(aByte(lngPosition + 4 + 4 + lng1)), lng2
    End If
    
    ' Unlock
    ' Automation error if uncommented !
    'dsb.Unlock ptr1, lng1, ptr2, lng2
    Exit Sub
End Sub
Private Sub picAnyTriple_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then Call MoveChip(Button, X, Y, imgAnyTriple, imgAnyTripleBet(DragChip), picAnyTriple.Tag)
    Exit Sub
End Sub
Private Sub picBig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then Call MoveChip(Button, X, Y, imgBig, imgBigBet(DragChip), picBig.Tag)
    Exit Sub
End Sub
Private Sub picCombi_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then
        Select Case Index
            Case 12
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiOneAndTwoBet(DragChip), picCombi(Index).Tag)
            Case 13
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiOneAndThreeBet(DragChip), picCombi(Index).Tag)
            Case 14
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiOneAndFourBet(DragChip), picCombi(Index).Tag)
            Case 15
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiOneAndFiveBet(DragChip), picCombi(Index).Tag)
            Case 16
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiOneAndSixBet(DragChip), picCombi(Index).Tag)
            Case 23
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiTwoAndThreeBet(DragChip), picCombi(Index).Tag)
            Case 24
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiTwoAndFourBet(DragChip), picCombi(Index).Tag)
            Case 25
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiTwoAndFiveBet(DragChip), picCombi(Index).Tag)
            Case 26
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiTwoAndSixBet(DragChip), picCombi(Index).Tag)
            Case 34
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiThreeAndFourBet(DragChip), picCombi(Index).Tag)
            Case 35
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiThreeAndFiveBet(DragChip), picCombi(Index).Tag)
            Case 36
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiThreeAndSixBet(DragChip), picCombi(Index).Tag)
            Case 45
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiFourAndFiveBet(DragChip), picCombi(Index).Tag)
            Case 46
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiFourAndSixBet(DragChip), picCombi(Index).Tag)
            Case 56
                Call MoveChip(Button, X, Y, imgCombi(Index), imgCombiFiveAndSixBet(DragChip), picCombi(Index).Tag)
        End Select
    End If
    Exit Sub
End Sub
Private Sub picCount_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then
    Select Case Index
        Case 4
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountFourBet(DragChip), picCount(Index).Tag)
        Case 5
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountFiveBet(DragChip), picCount(Index).Tag)
        Case 6
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountSixBet(DragChip), picCount(Index).Tag)
        Case 7
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountSevenBet(DragChip), picCount(Index).Tag)
        Case 8
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountEightBet(DragChip), picCount(Index).Tag)
        Case 9
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountNineBet(DragChip), picCount(Index).Tag)
        Case 10
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountTenBet(DragChip), picCount(Index).Tag)
        Case 11
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountElevenBet(DragChip), picCount(Index).Tag)
        Case 12
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountTwelveBet(DragChip), picCount(Index).Tag)
        Case 13
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountThirteenBet(DragChip), picCount(Index).Tag)
        Case 14
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountFourteenBet(DragChip), picCount(Index).Tag)
        Case 15
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountFifthteenBet(DragChip), picCount(Index).Tag)
        Case 16
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountSixteenBet(DragChip), picCount(Index).Tag)
        Case 17
            Call MoveChip(Button, X, Y, imgCount(Index), imgCountSeventeenBet(DragChip), picCount(Index).Tag)
        End Select
    End If
    Exit Sub
End Sub
Private Sub picDouble_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then
    Select Case Index
        Case 1
            Call MoveChip(Button, X, Y, imgDouble(Index), imgDoubleOneBet(DragChip), picDouble(Index).Tag)
        Case 2
            Call MoveChip(Button, X, Y, imgDouble(Index), imgDoubleTwoBet(DragChip), picDouble(Index).Tag)
        Case 3
            Call MoveChip(Button, X, Y, imgDouble(Index), imgDoubleThreeBet(DragChip), picDouble(Index).Tag)
        Case 4
            Call MoveChip(Button, X, Y, imgDouble(Index), imgDoubleFourBet(DragChip), picDouble(Index).Tag)
        Case 5
            Call MoveChip(Button, X, Y, imgDouble(Index), imgDoubleFiveBet(DragChip), picDouble(Index).Tag)
        Case 6
            Call MoveChip(Button, X, Y, imgDouble(Index), imgDoubleSixBet(DragChip), picDouble(Index).Tag)
        End Select
    End If
    Exit Sub
End Sub
Private Sub picSingle_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then
        Select Case Index
            Case 1
                Call MoveChip(Button, X, Y, imgSingle(Index), imgSingleOneBet(DragChip), picSingle(Index).Tag)
            Case 2
                Call MoveChip(Button, X, Y, imgSingle(Index), imgSingleTwoBet(DragChip), picSingle(Index).Tag)
            Case 3
                Call MoveChip(Button, X, Y, imgSingle(Index), imgSingleThreeBet(DragChip), picSingle(Index).Tag)
            Case 4
                Call MoveChip(Button, X, Y, imgSingle(Index), imgSingleFourBet(DragChip), picSingle(Index).Tag)
            Case 5
                Call MoveChip(Button, X, Y, imgSingle(Index), imgSingleFiveBet(DragChip), picSingle(Index).Tag)
            Case 6
                Call MoveChip(Button, X, Y, imgSingle(Index), imgSingleSixBet(DragChip), picSingle(Index).Tag)
        End Select
    End If
    Exit Sub
End Sub
Private Sub picSmall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then Call MoveChip(Button, X, Y, imgSmall, imgSmallBet(DragChip), picSmall.Tag)
    Exit Sub
End Sub
Private Sub picTriple_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.MousePointer = 99 Then
    Select Case Index
        Case 1
            Call MoveChip(Button, X, Y, imgTriple(Index), imgTripleOneBet(DragChip), picTriple(Index).Tag)
        Case 2
            Call MoveChip(Button, X, Y, imgTriple(Index), imgTripleTwoBet(DragChip), picTriple(Index).Tag)
        Case 3
            Call MoveChip(Button, X, Y, imgTriple(Index), imgTripleThreeBet(DragChip), picTriple(Index).Tag)
        Case 4
            Call MoveChip(Button, X, Y, imgTriple(Index), imgTripleFourBet(DragChip), picTriple(Index).Tag)
        Case 5
            Call MoveChip(Button, X, Y, imgTriple(Index), imgTripleFiveBet(DragChip), picTriple(Index).Tag)
        Case 6
            Call MoveChip(Button, X, Y, imgTriple(Index), imgTripleSixBet(DragChip), picTriple(Index).Tag)
        End Select
    End If
    Exit Sub
End Sub
Private Sub picChip_Click(Index As Integer)
    Reset
    If Game.CreditsLeft - Key2Value(CByte(Index)) < 0 Then
        Beep
    Else
        Me.MousePointer = 99
        Me.MouseIcon = ImageList1(Index).ListImages(1).Picture
        DragChip = Index
    End If
    Exit Sub
End Sub
