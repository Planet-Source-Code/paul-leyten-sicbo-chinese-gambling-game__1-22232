VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"bet"
Attribute VB_Ext_KEY = "Member1" ,"BetType"
Attribute VB_Ext_KEY = "Member2" ,"colImg"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarBet As Double 'local copy
Private mvarsKey As Variant 'local copy
Private mvarBetType As Integer 'local copy
Private mvarcolImg As colImg
'local variable(s) to hold property value(s)
Private mvarPriceWinning As Boolean 'local copy
Public Sub ResetPrice()
    mvarPriceWinning = False
    Exit Sub
End Sub
Public Sub Win()
    mvarPriceWinning = True
    Exit Sub
End Sub
Public Property Get PriceWinning() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PriceWinning
    PriceWinning = mvarPriceWinning
    Exit Property
End Property
Public Property Get colImg() As colImg
    If mvarcolImg Is Nothing Then
        Set mvarcolImg = New colImg
    End If
    Set colImg = mvarcolImg
    Exit Property
End Property
Public Property Set colImg(vData As colImg)
    Set mvarcolImg = vData
    Exit Property
End Property
Private Sub Class_Terminate()
    Set mvarcolImg = Nothing
    Exit Sub
End Sub
Public Property Let BetType(ByVal vData As Integer)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.BetType = Form1
    mvarBetType = vData
    Exit Property
End Property
Public Property Get BetType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BetType
    BetType = mvarBetType
    Exit Property
End Property
Public Property Let sKey(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sKey = 5
    mvarsKey = vData
    Exit Property
End Property
Public Property Get sKey() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.sKey
    If IsObject(mvarsKey) Then
        Set sKey = mvarsKey
    Else
        sKey = mvarsKey
    End If
    Exit Property
End Property
Public Property Let bet(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Bet = 5
    mvarBet = vData
    Exit Property
End Property
Public Property Get bet() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Bet
    bet = mvarBet
    Exit Property
End Property
