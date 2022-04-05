VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGlobalFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mRgx As VBScript_RegExp_55.RegExp


Public Sub DisplayErrMessage(ByVal pClass As String, ByVal pProcedure As String)

SetVariables:
    Dim msg As String
    Let msg = Err.Source & vbNewLine & Err.Description

DisplayUserMessage:
    Call MsgBox(msg, vbCritical, "Class " & pClass & "\" & pProcedure)

ShowDebugMessage:
    Debug.Print String(50, "=")
    Debug.Print "ClassName: " & pClass
    Debug.Print "Procedure: " & pProcedure
    Debug.Print "ErrSource: " & Err.Source
    Debug.Print "ErrNumber: " & Abs(Err.Number)
    Debug.Print "ErrDescription: " & Err.Description
    Debug.Print "ErrDate: " & Now()
    Debug.Print String(50, "=") & vbNewLine
    
End Sub


Public Function StringToMD5Hex(ByVal pString As String) As String
   
SetVariables:
    Dim enc As Object
    Dim bytes() As Byte
    Dim pos As Long
    Dim outstr As String

CreateEncoderObject:
    On Error GoTo ErrHandler
    
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

GetBytes:
    Let bytes = StrConv(pString, vbFromUnicode)
    Let bytes = enc.ComputeHash_2(bytes)

CreateStringMD5:
    For pos = LBound(bytes) To UBound(bytes)
        Let outstr = outstr & LCase(Right("0" & Hex(bytes(pos)), 2))
    Next pos

ReturningValue:
    Let StringToMD5Hex = outstr
    
    On Error GoTo 0

Finish:
    On Error Resume Next
    Set enc = Nothing

ErrHandler:
    Call DisplayErrMessage("GlobalFunctions", "StringToMD5Hex")
    GoTo Finish

End Function


Public Function RegexFormat(ByVal pTextString As String, ByVal pReplaceNumber As Boolean) As String

SetVariables:
    Set mRgx = New RegExp
    Let mRgx.Global = True
    Let mRgx.MultiLine = True

ReplaceDiacritics:
    Let mRgx.Pattern = "[‡·‰‚„]"
    Let pTextString = mRgx.Replace(pTextString, "a")
    Let mRgx.Pattern = "[¿¡ƒ¬√]"
    Let pTextString = mRgx.Replace(pTextString, "A")
    Let mRgx.Pattern = "[ËÈÎÍ]"
    Let pTextString = mRgx.Replace(pTextString, "e")
    Let mRgx.Pattern = "[»…À ]"
    Let pTextString = mRgx.Replace(pTextString, "E")
    Let mRgx.Pattern = "[ÏÌÔÓ]"
    Let pTextString = mRgx.Replace(pTextString, "mCount")
    Let mRgx.Pattern = "[ÃÕœŒ]"
    Let pTextString = mRgx.Replace(pTextString, "I")
    Let mRgx.Pattern = "[ÚÛˆıÙ]"
    Let pTextString = mRgx.Replace(pTextString, "o")
    Let mRgx.Pattern = "[“”÷’‘]"
    Let pTextString = mRgx.Replace(pTextString, "O")
    Let mRgx.Pattern = "[˘˙¸˚]"
    Let pTextString = mRgx.Replace(pTextString, "u")
    Let mRgx.Pattern = "[Ÿ⁄‹€]"
    Let pTextString = mRgx.Replace(pTextString, "U")
    Let mRgx.Pattern = "[Á]"
    Let pTextString = mRgx.Replace(pTextString, "c")
    Let mRgx.Pattern = "[«]"
    Let pTextString = mRgx.Replace(pTextString, "C")
    Let mRgx.Pattern = "[Ò]"
    Let pTextString = mRgx.Replace(pTextString, "n")
    Let mRgx.Pattern = "[—]"
    Let pTextString = mRgx.Replace(pTextString, "N")

ReplaceMultipleSpaces:
    Let mRgx.Pattern = "\s{2,}"
    Let pTextString = mRgx.Replace(pTextString, Space(1))

ReplaceSymbols:
    Let mRgx.Pattern = "[^a-zA-Z0-9\s]"
    Let pTextString = mRgx.Replace(pTextString, vbNullString)

ReplaceNumbers:
    If (Not pReplaceNumber) Then GoTo TrimText
    Let mRgx.Pattern = "[^a-zA-Z\s]"
    Let pTextString = mRgx.Replace(pTextString, vbNullString)

TrimText:
    Let RegexFormat = Trim(pTextString)
    Let RegexFormat = StrConv(RegexFormat, vbUpperCase)
    
Finish:
    On Error Resume Next
    Exit Function

ErrHandler:
   Call DisplayErrMessage("GlobalFunctions", "RegexFormat")
   GoTo Finish
    
End Function


Public Function CurrencyFormat(ByVal pCurrency As String) As String
    
SetVariables:
    Let pCurrency = Replace(pCurrency, ",", ".")

ReturningValue:
    Let CurrencyFormat = pCurrency

Finish:
    On Error Resume Next
    Exit Function

ErrHandler:
    Call DisplayErrMessage("GlobalFunctions", "CurrencyFormat")
    GoTo Finish
    
End Function


Public Function HoursToMinutes(ByVal pHours As String) As String
    
SetVariables:
    Dim h As Integer
    Dim m As Integer
    
    On Error GoTo ErrHandler
    
    Let HoursToMinutes = 0
    
    If (Not IsDate(pHours)) Then
        GoTo Finish
    End If

CalculateHourToMinutes:
    Let h = Split(pHours, ":")(0)
    Let m = Split(pHours, ":")(1)
    
GerarValor:
    Let HoursToMinutes = (h * 60) + m
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Exit Function
    
ErrHandler:
    Call DisplayErrMessage("GlobalFunctions", "HoursToMinutes")
    GoTo Finish
    
End Function


Public Function MinutesToHours(ByVal pMinutes As Long) As String
    
SetVariables:
    Dim h As Integer
    Dim m As Integer
    
    On Error GoTo ErrHandler
    
    Let MinutesToHours = "00:00"
    
    If (pMinutes = 0) Then
        Exit Function
    End If

CalculateVariables:
    Let h = pMinutes / 60
    Let m = pMinutes Mod 60
    
ReturningValue:
    Let MinutesToHours = Format(h, "00:") & Format(m, "00")
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Exit Function
    
ErrHandler:
    Call DisplayErrMessage("GlobalFunctions", "MinutesToHours")
    GoTo Finish
    
End Function
