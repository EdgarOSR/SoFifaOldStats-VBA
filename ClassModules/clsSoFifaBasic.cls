VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoFifaBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const mImg As String = "https://cdn.sofifa.net/mayers/notfound_0_120.png"
Private mName As String
Private mNation As String
Private mBirth As String
Private mFoot As String
Private mPrefPos As String
Private mPositions As String
Private mHeight As String
Private mWeight As String
Private mOverall As String
Private mWeakFoot As String
Private mSkillMoves As String
Private mInterRep As String
Private oDict As Dictionary


Public Property Get GetDictionary() As Dictionary
      
      Set oDict = New Dictionary
      
      With oDict
        Call .Add("name", mName)
        Call .Add("imageUrl", mImg)
        Call .Add("positions", mPositions)
        Call .Add("nationality", mNation)
        Call .Add("birth_date", mBirth)
        Call .Add("height", mHeight)
        Call .Add("weight", mWeight)
        Call .Add("overall_rating", mOverall)
        Call .Add("preferred_foot", mFoot)
        Call .Add("weak_foot", mWeakFoot)
        Call .Add("skill_moves", mSkillMoves)
        Call .Add("international_reputation", mInterRep)
        Call .Add("position_prefered_team", mPrefPos)
      End With

      Set GetDictionary = oDict
      
End Property


Public Property Let Name(ByVal vNewValue As String)
      Let mName = Trim(vNewValue)
End Property


Public Property Let Nation(ByVal vNewValue As String)
      mNation = StrConv(Trim(vNewValue), vbProperCase)
End Property


Public Property Let Basics(ByVal vNewValue As String)
    Dim aux As String
    Let aux = Split(vNewValue, ")")(0)
    Let mBirth = Right(aux, 4)
    Let mBirth = Val(mBirth) + (Year(Date) - 2006)
    Let aux = Split(vNewValue, ") ")(1)
    Let mHeight = Val(Split(aux)(0))
    Let mWeight = Val(Split(aux)(1))
End Property


Public Property Let Overall(ByVal vNewValue As String)
      Let mOverall = Val(vNewValue)
End Property


Public Property Let Foot(ByVal vNewValue As String)
      Let mFoot = Trim(vNewValue)
End Property


Public Property Let WeakFoot(ByVal vNewValue As String)
      Let mWeakFoot = Val(vNewValue)
End Property


Public Property Let SkillMoves(ByVal vNewValue As String)
      Let mSkillMoves = Val(vNewValue)
End Property


Public Property Let InterRep(ByVal vNewValue As String)
      Let mInterRep = Val(vNewValue)
End Property


Public Property Let PrefPos(ByVal vNewValue As String)
      Let mPrefPos = Trim(vNewValue)
End Property


Public Property Let Positions(ByVal vNewValue As String)
    'output: ["ST","ST"]
    Let mPositions = Chr(91) & Chr(34) & mPrefPos & Chr(34) & Chr(44)
    Let mPositions = mPositions & Chr(34) & Trim(vNewValue) & Chr(34) & Chr(93)
End Property
