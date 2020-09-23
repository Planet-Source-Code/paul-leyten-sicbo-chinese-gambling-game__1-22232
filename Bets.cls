VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"bet"
Attribute VB_Ext_KEY = "Member1" ,"BetType"
Attribute VB_Ext_KEY = "Member2" ,"colBets"
Option Explicit
Public Event Win(Price As Double)
Private mvarBetTotal As Double
Private mvarCreditsLeft As Double 'local copy
Private mvarForm As Form 'local copy
Private mvarBet As bet
Private mvarWinPrice As Double
Private mvarcolBets As colBets
Private mvarBetType As BetType
Public Property Let CreditsLeft(ByVal vData As Double)
    mvarCreditsLeft = vData
    Exit Property
End Property
Public Property Get BetType() As BetType
    If mvarBetType Is Nothing Then
        Set mvarBetType = New BetType
    End If
    Set BetType = mvarBetType
    Exit Property
End Property
Public Property Set BetType(vData As BetType)
    Set mvarBetType = vData
    Exit Property
End Property
Public Property Get colBets() As colBets
    If mvarcolBets Is Nothing Then
        Set mvarcolBets = New colBets
    End If
    Set colBets = mvarcolBets
    Exit Property
End Property
Public Property Set colBets(vData As colBets)
    Set mvarcolBets = vData
    Exit Property
End Property
Public Property Get bet() As bet
    If mvarBet Is Nothing Then
        Set mvarBet = New bet
    End If
    Set bet = mvarBet
    Exit Property
End Property
Public Property Set bet(vData As bet)
    Set mvarBet = vData
    Exit Property
End Property
Private Sub Class_Terminate()
  Set mvarBetType = Nothing
  Set mvarcolBets = Nothing
    Set mvarBet = Nothing
    Exit Sub
End Sub
Public Sub Add2Bet(bet As Double)
    mvarCreditsLeft = mvarCreditsLeft - bet
    mvarBetTotal = mvarBetTotal + bet
    mvarForm.txtBet.Text = mvarBetTotal
    mvarForm.txtCredits.Text = mvarCreditsLeft
    Exit Sub
End Sub
Public Property Get CurrentBet() As Double
    CurrentBet = mvarBetTotal
    Exit Property
End Property
Public Sub StartGame()
    mvarForm.txtCredits.Text = Format(mvarCreditsLeft, "0,00")
    mvarForm.txtBet = ""
    mvarForm.txtWin = ""
    Exit Sub
End Sub
Public Property Set Form(ByVal vData As Form)
    Set mvarForm = vData
    Exit Property
End Property
Public Property Get CreditsLeft() As Double
    mvarForm.txtCredits.Text = mvarCreditsLeft
    CreditsLeft = mvarCreditsLeft
    Exit Property
End Property
Public Function CheckWins(D1 As Byte, D2 As Byte, D3 As Byte) As Long
Dim iCount, jCount As Integer
Dim blnStop As Boolean
Dim xCount As Integer
Call Lights(D1, D2, D3)
For jCount = 1 To mvarcolBets.Count
    mvarWinPrice = mvarWinPrice + Prices(mvarcolBets, jCount, D1, D2, D3)
Next
iCount = mvarcolBets.Count
While Not blnStop
  If Not mvarcolBets.item(iCount).PriceWinning Then
    mvarBetTotal = mvarBetTotal - mvarcolBets.item(iCount).bet
    mvarcolBets.Remove iCount
    iCount = mvarcolBets.Count
  Else
    iCount = iCount - 1
  End If
  If iCount <= 0 Then blnStop = True
Wend
    
For iCount = 1 To mvarcolBets.Count
    mvarcolBets.item(iCount).ResetPrice
Next
CheckWins = mvarWinPrice
mvarForm.txtWin = mvarWinPrice
mvarForm.txtBet = mvarBetTotal
mvarWinPrice = 0
    Exit Function
End Function
Private Function Prices(cb As colBets, iCount As Integer, D1 As Byte, D3 As Byte, D2 As Byte) As Double
Dim intResult As Byte
intResult = D1 + D2 + D3
Dim dblbet As Double
dblbet = cb.item(iCount).bet
    Select Case cb.item(iCount).sKey
        Case "imgSmall"
            If intResult > 3 And intResult < 10 Then
                Prices = (1 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgBig"
            If intResult > 10 And intResult < 18 Then
                Prices = (1 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgDouble1"
            If (D1 = 1 And D2 = 1) Or (D2 = 1 And D3 = 1) Or (D1 = 1 And D3 = 1) Then
                Prices = (10 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgDouble2"
            If (D1 = 2 And D2 = 2) Or (D2 = 2 And D3 = 2) Or (D1 = 2 And D3 = 2) Then
                Prices = (10 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgDouble3"
            If (D1 = 3 And D2 = 3) Or (D2 = 3 And D3 = 3) Or (D1 = 3 And D3 = 3) Then
                Prices = (10 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgDouble4"
            If (D1 = 4 And D2 = 4) Or (D2 = 4 And D3 = 4) Or (D1 = 4 And D3 = 4) Then
                Prices = (10 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgDouble5"
            If (D1 = 5 And D2 = 5) Or (D2 = 5 And D3 = 5) Or (D1 = 5 And D3 = 5) Then
                Prices = (10 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgDouble6"
            If (D1 = 6 And D2 = 6) Or (D2 = 6 And D3 = 6) Or (D1 = 6 And D3 = 6) Then
                Prices = (10 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgAnyTriple"
            If D1 = D2 And D2 = D3 And D3 = D1 Then
                Prices = (30 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgTriple1"
            If D1 = 1 And D2 = 1 And D3 = 1 Then
                Prices = (150 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgTriple2"
            If D1 = 2 And D2 = 2 And D3 = 2 Then
                Prices = (150 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgTriple3"
            If D1 = 3 And D2 = 2 And D3 = 3 Then
                Prices = (150 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgTriple4"
            If D1 = 4 And D2 = 2 And D3 = 4 Then
                Prices = (150 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgTriple5"
            If D1 = 5 And D2 = 2 And D3 = 5 Then
                Prices = (150 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgTriple6"
            If D1 = 6 And D2 = 2 And D3 = 6 Then
                Prices = (150 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount4"
            If intResult = 4 Then
                Prices = (60 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount5"
            If intResult = 5 Then
                Prices = (18 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount6"
            If intResult = 6 Then
                Prices = (12 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount7"
            If intResult = 7 Then
                Prices = (8 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount8"
            If intResult = 8 Then
                Prices = (6 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount9"
            If intResult = 9 Then
                Prices = (6 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount10"
            If intResult = 10 Then
                Prices = (6 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount11"
            If intResult = 11 Then
                Prices = (6 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount12"
            If intResult = 12 Then
                Prices = (6 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount13"
            If intResult = 13 Then
                Prices = (8 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount14"
            If intResult = 14 Then
                Prices = (12 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount15"
            If intResult = 15 Then
                Prices = (14 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount16"
            If intResult = 16 Then
                Prices = (18 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCount17"
            If intResult = 17 Then
                Prices = (60 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi12"
            If (D1 = 1 Or D2 = 1 Or D3 = 1) And (D1 = 2 Or D2 = 2 Or D3 = 2) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi13"
            If (D1 = 1 Or D2 = 1 Or D3 = 1) And (D1 = 3 Or D2 = 3 Or D3 = 3) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi14"
            If (D1 = 1 Or D2 = 1 Or D3 = 1) And (D1 = 4 Or D2 = 4 Or D3 = 4) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi15"
            If (D1 = 1 Or D2 = 1 Or D3 = 1) And (D1 = 5 Or D2 = 5 Or D3 = 5) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi16"
            If (D1 = 1 Or D2 = 1 Or D3 = 1) And (D1 = 6 Or D2 = 6 Or D3 = 6) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi23"
            If (D1 = 2 Or D2 = 2 Or D3 = 2) And (D1 = 3 Or D2 = 3 Or D3 = 3) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi24"
            If (D1 = 2 Or D2 = 2 Or D3 = 2) And (D1 = 4 Or D2 = 4 Or D3 = 4) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi25"
            If (D1 = 2 Or D2 = 2 Or D3 = 2) And (D1 = 5 Or D2 = 5 Or D3 = 5) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi26"
            If (D1 = 2 Or D2 = 2 Or D3 = 2) And (D1 = 6 Or D2 = 6 Or D3 = 6) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi34"
            If (D1 = 3 Or D2 = 3 Or D3 = 3) And (D1 = 4 Or D2 = 4 Or D3 = 4) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi35"
            If (D1 = 3 Or D2 = 3 Or D3 = 3) And (D1 = 5 Or D2 = 5 Or D3 = 5) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi36"
            If (D1 = 3 Or D2 = 3 Or D3 = 3) And (D1 = 6 Or D2 = 6 Or D3 = 6) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi45"
            If (D1 = 4 Or D2 = 4 Or D3 = 4) And (D1 = 5 Or D2 = 5 Or D3 = 5) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi46"
            If (D1 = 4 Or D2 = 4 Or D3 = 4) And (D1 = 6 Or D2 = 6 Or D3 = 6) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgCombi56"
            If (D1 = 5 Or D2 = 5 Or D3 = 5) And (D1 = 6 Or D2 = 6 Or D3 = 6) Then
                Prices = (5 * dblbet)
                cb.item(iCount).Win
            End If
        Case "imgSingle1"
            If D1 = 1 Or D2 = 1 Or D3 = 1 Then
                Prices = SingleBet(1, D1, D3, D2, dblbet)
                cb.item(iCount).Win
            End If
        Case "imgSingle2"
            If D1 = 2 Or D2 = 2 Or D3 = 2 Then
                Prices = SingleBet(2, D1, D3, D2, dblbet)
                cb.item(iCount).Win
            End If
        Case "imgSingle3"
            If D1 = 3 Or D2 = 3 Or D3 = 3 Then
                Prices = SingleBet(3, D1, D3, D2, dblbet)
                cb.item(iCount).Win
            End If
        Case "imgSingle4"
            If D1 = 4 Or D2 = 4 Or D3 = 4 Then
                Prices = SingleBet(4, D1, D3, D2, dblbet)
                cb.item(iCount).Win
            End If
        Case "imgSingle5"
            If D1 = 5 Or D2 = 5 Or D3 = 5 Then
                Prices = SingleBet(5, D1, D3, D2, dblbet)
                cb.item(iCount).Win
            End If
        Case "imgSingle6"
            If D1 = 6 Or D2 = 6 Or D3 = 6 Then
                Prices = SingleBet(6, D1, D3, D2, dblbet)
                cb.item(iCount).Win
            End If
        Case Else
            Err.Raise 666, , "Duivels!, onmogelijk"
    End Select
    Exit Function
End Function
Private Function SingleBet(ref As Byte, D1 As Byte, D3 As Byte, D2 As Byte, dblbet As Double) As Double
    If (D1 = ref And D2 = ref) And (D2 = ref And D3 = ref) Then ' tripple!
        SingleBet = (3 * dblbet)
    ElseIf (D1 = ref And D2 = ref) Or (D1 = ref And D3 = ref) Or (D2 = ref And D3 = ref) Then ' double!
        SingleBet = (2 * dblbet)
    Else
        SingleBet = (1 * dblbet)
    End If
    Exit Function
End Function
Private Sub Lights(D1 As Byte, D3 As Byte, D2 As Byte)
    Dim intResult As Integer
    
    DoEvents
    mvarForm.imgSingle(D1).Visible = True
    DoEvents
    mvarForm.imgSingle(D2).Visible = True
    DoEvents
    mvarForm.imgSingle(D3).Visible = True
    
    On Error Resume Next ' Er zijn ON-mogelijkheden
    If D1 < D2 Then
        mvarForm.imgCombi(D1 & D2).Visible = True
    Else
        mvarForm.imgCombi(D2 & D1).Visible = True
    End If
    If D1 < D3 Then
        mvarForm.imgCombi(D1 & D3).Visible = True
    Else
        mvarForm.imgCombi(D3 & D1).Visible = True
    End If
    If D2 < D3 Then
        mvarForm.imgCombi(D2 & D3).Visible = True
    Else
        mvarForm.imgCombi(D3 & D2).Visible = True
    End If
    
    'Dices Count+Small/Big
    intResult = D1 + D2 + D3
    If intResult > 3 And intResult < 18 Then
        mvarForm.imgCount(intResult).Visible = True
        If intResult < 11 Then
            mvarForm.imgSmall.Visible = True
        Else
            mvarForm.imgBig.Visible = True
        End If
    End If
    If D1 = D2 Then
        mvarForm.imgDouble(D1).Visible = True
        mvarForm.imgDouble(D2).Visible = True
    ElseIf D2 = D3 Then
        mvarForm.imgDouble(D2).Visible = True
        mvarForm.imgDouble(D3).Visible = True
    ElseIf D1 = D3 Then
        mvarForm.imgDouble(D1).Visible = True
        mvarForm.imgDouble(D3).Visible = True
    End If
    If D1 = D2 And D2 = D3 Then ' Triple!, Reset Big/Small if tripple
        mvarForm.imgTriple(D1).Visible = True
        mvarForm.imgAnyTriple.Visible = True
        mvarForm.imgBig.Visible = False
        mvarForm.imgSmall.Visible = False
    End If
    Exit Sub
End Sub
