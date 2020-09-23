VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Bet"
Attribute VB_Ext_KEY = "Member0" ,"Bet"
Option Explicit
'local variable to hold collection
Private mCol As Collection
Public Function Add(bet As Integer, imgLeft As Single, imgTop As Single, cImage As Image, sKey As String) As bet
    'create a new object
    Dim objNewMember As bet
    Set objNewMember = New bet
    'set the properties passed into the method
    objNewMember.bet = bet
    objNewMember.colImg.Add imgLeft, imgTop, cImage, bet2sKey(bet)
    objNewMember.sKey = sKey
   
    mCol.Add objNewMember, sKey
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
End Function
Public Property Get item(vntIndexKey As Variant) As bet
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    
  Set item = mCol(vntIndexKey)
    Exit Property
End Property
Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
    Exit Property
End Property
Public Sub Remove(vntIndexKey As Variant)
Dim iCount As Integer
'remove images!
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    Dim objMember As bet
    Set objMember = New bet
    Set objMember = mCol.item(vntIndexKey)
    
    For iCount = 1 To objMember.colImg.Count
        objMember.colImg(iCount).cImage.Picture = LoadPicture()
        objMember.colImg(iCount).cImage.Left = 0
        objMember.colImg(iCount).cImage.Top = 0
        objMember.colImg(iCount).cImage.Visible = False
    Next
    
    mCol.Remove vntIndexKey
    Exit Sub
End Sub
Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
    Exit Property
End Property
Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Exit Sub
End Sub
Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Exit Sub
End Sub
