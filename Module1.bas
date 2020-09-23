Attribute VB_Name = "Module1"
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
Public Declare Function lstrcpy Lib "Kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Function bet2sKey(bet As Integer) As String
    Select Case bet
        Case 1
            bet2sKey = "One"
        Case 5
            bet2sKey = "Five"
        Case 25
            bet2sKey = "TwentyFive"
        Case 100
            bet2sKey = "Hundred"
        Case 500
            bet2sKey = "FiveHundred"
    End Select
        
    Exit Function
End Function
Function Key2Value(Idx As Byte) As Integer
Select Case Idx
    Case 1
         Key2Value = 1
    Case 2
        Key2Value = 5
    Case 3
        Key2Value = 25
    Case 4
        Key2Value = 100
    Case 5
        Key2Value = 500
End Select
    Exit Function
End Function
Function Value2Key(varVal As Integer) As String
Select Case Idx
    Case 1
         Value2Key = "One"
    Case 2
         Value2Key = "Five"
    Case 3
         Value2Key = "TwentyFive"
    Case 4
         Value2Key = "Hundred"
    Case 5
         Value2Key = "FiveHundred"
End Select
    Exit Function
End Function
Function GetIndex(Idx As Object) As String
  On Error GoTo Err_GetIndex
  GetIndex = CStr(Idx.Index)
  Exit Function
  
Err_GetIndex:
  GetIndex = ""
    Exit Function
End Function
