VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParseJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private oClsGF As clsGlobalFunc
Private mJSON As String


Public Function GetJSON(ByVal pDicBasics As Scripting.Dictionary, ByVal pDicSkills As Scripting.Dictionary) As String

Variables:
    Set oClsGF = New clsGlobalFunc
    Let mJSON = vbNullString

ParsingJSON:
    Call ParseBasics(pDicBasics)
    Call ParseSkills(pDicSkills)

'    Debug.Print mJSON
    
ReturningValue:
    Let GetJSON = mJSON
    
Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Exit Function
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("ParseJSON", "GetJSON")
    GoTo Finish

End Function


Private Sub ParseBasics(ByVal pDic As Scripting.Dictionary)

VariablesDim:
    Dim vKey As Variant
    Dim vJSON As String
    
VariablesSet:
    Set oClsGF = New clsGlobalFunc
    
CreateJSON:
    Let vJSON = vbNullString
    Let vJSON = Chr(123)
    
    For Each vKey In pDic.Keys
        If (IsEmpty(vKey)) Then
            Exit For
        ElseIf (vKey = "positions") Then
            Let vJSON = vJSON & Chr(34) & vKey & Chr(34) & Chr(58) & pDic(vKey)
        ElseIf (IsNumeric(pDic(vKey))) Then
            Let vJSON = vJSON & Chr(34) & vKey & Chr(34) & Chr(58) & pDic(vKey)
        Else
            Let vJSON = vJSON & Chr(34) & vKey & Chr(34) & Chr(58) & Chr(34) & pDic(vKey) & Chr(34)
        End If
        Let vJSON = vJSON & Chr(44)
    Next vKey
    
    Let mJSON = mJSON & vJSON

Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set vKey = Nothing
    Exit Sub
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("ParseJSON", "ParseBasics")
    GoTo Finish

End Sub

Private Sub ParseSkills(ByVal pDic As Scripting.Dictionary)

VariablesDim:
    Dim vKey As Variant
    Dim vJSON As String
    
VariablesSet:
    Set oClsGF = New clsGlobalFunc
    
CreateJSON:
    Let vJSON = vbNullString
    Let vJSON = Chr(34) & "skills" & Chr(34) & Chr(58) & Chr(123)
    
    For Each vKey In pDic.Keys
        If (IsEmpty(vKey)) Then
            Exit For
        ElseIf (vKey = "player_traits") Then
            Let vJSON = Left(vJSON, Len(vJSON) - 1)
            Let vJSON = vJSON & Chr(125) & Chr(44)
            Let vJSON = vJSON & Chr(34) & vKey & Chr(34) & Chr(58) & pDic(vKey) & Chr(125)
            Exit For
        ElseIf (IsNumeric(pDic(vKey))) Then
            Let vJSON = vJSON & Chr(34) & vKey & Chr(34) & Chr(58) & pDic(vKey)
        Else
            Let vJSON = vJSON & Chr(34) & vKey & Chr(34) & Chr(58) & Chr(34) & pDic(vKey) & Chr(34)
        End If
        Let vJSON = vJSON & Chr(44)
    Next vKey
    
    Let mJSON = mJSON & vJSON

Finish:
    On Error Resume Next
    Set oClsGF = Nothing
    Set vKey = Nothing
    Exit Sub
    
ErrHandler:
    Call oClsGF.DisplayErrMessage("ParseJSON", "ParseSkills")
    GoTo Finish

End Sub

