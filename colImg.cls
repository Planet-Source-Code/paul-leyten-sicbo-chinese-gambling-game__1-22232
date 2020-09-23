VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"BetType"
Attribute VB_Ext_KEY = "Member0" ,"BetType"
'local variable to hold collection
Private mCol As Collection
Public Function Add(imgLeft As Single, imgTop As Single, cImage As Image, sKey As Variant) As BetType
    'create a new object
    Dim objNewMember As BetType
    Set objNewMember = New BetType
    'set the properties passed into the method
    objNewMember.imgLeft = imgLeft
    objNewMember.imgTop = imgTop
    objNewMember.AddChip
    If IsObject(cImage) Then
        Set objNewMember.cImage = cImage
    Else
        objNewMember.cImage = cImage
    End If
    objNewMember.sKey = sKey
    mCol.Add objNewMember, sKey
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
    Exit Function
End Function
Public Property Get item(vntIndexKey As Variant) As BetType
Attribute item.VB_UserMemId = 0
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
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
    Exit Sub
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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
