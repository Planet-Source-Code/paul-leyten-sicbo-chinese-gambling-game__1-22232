VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BetType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"BetType"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'local variable(s) to hold property value(s)
Private mvarimgLeft As Integer 'local copy
Private mvarimgTop As Integer 'local copy
Private mvarcImage As Image 'local copy
Private mvarBetType As Integer 'local copy
Private mvarsKey As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarCountChips As Integer 'local copy
Public Sub RemoveImg()
    
    mvarcImage.Visible = False
    Set mvarcImage = Nothing
    'Hierook het plaatje verwijderen!
    Exit Sub
End Sub
Public Property Get CountChips() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CountChips
    CountChips = mvarCountChips
    Exit Property
End Property
'local variable(s) to hold property value(s)
Public Sub AddChip()
    mvarCountChips = mvarCountChips + 1
    Exit Sub
End Sub
Public Property Let sKey(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.sKey = 5
    mvarsKey = vData
    Exit Property
End Property
Public Property Set sKey(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.sKey = Form1
    Set mvarsKey = vData
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
Public Property Set cImage(ByVal vData As Image)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.cImage = Form1
    Set mvarcImage = vData
    Exit Property
End Property
Public Property Get cImage() As Image
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.cImage
    Set cImage = mvarcImage
    Exit Property
End Property
Public Property Let imgTop(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.imgTop = 5
    mvarimgTop = vData
    Exit Property
End Property
Public Property Get imgTop() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgTop
    imgTop = mvarimgTop
    Exit Property
End Property
Public Property Let imgLeft(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.imgLeft = 5
    mvarimgLeft = vData
    Exit Property
End Property
Public Property Get imgLeft() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.imgLeft
    imgLeft = mvarimgLeft
    Exit Property
End Property
