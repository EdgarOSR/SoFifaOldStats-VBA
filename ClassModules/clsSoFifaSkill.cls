VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOldStatsSkill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cgf As New clsGlobalFunc
Private oDic As Scripting.Dictionary
Private mCrossing As String
Private mFinishing As String
Private mHeadingAcc As String
Private mShortPass As String
Private mDribbling As String
Private mFkAcc As String
Private mLongPass As String
Private mBallControl As String
Private mAcceleration As String
Private mSprintSpeed As String
Private mReactions As String
Private mShotPower As String
Private mStamina As String
Private mStrength As String
Private mLongShots As String
Private mAggression As String
Private mInterceptions As String
Private mPositioning As String
Private mPenalties As String
Private mMarking As String
Private mStandingTackle As String
Private mGkDiving As String
Private mGkHandling As String
Private mGkKicking As String
Private mGkPositioning As String
Private mGkReflexes As String
Private mTraits As String

Private Property Get Volleys() As String
    Let Volleys = Val(mCrossing) + Val(mFinishing) + Val(mLongShots)
    Let Volleys = Round(Volleys / 3, 0)
End Property


Private Property Get Curve() As String
    Let Curve = Val(mFkAcc) + Val(mBallControl)
    Let Curve = Round(Curve / 2, 0)
End Property


Private Property Get Agility() As String
    Let Agility = Val(mAcceleration) + Val(mSprintSpeed) + Val(mReactions)
    Let Agility = Round(Agility / 3, 0)
End Property


Private Property Get Balance() As String
    Let Balance = Val(mStrength) + Val(mReactions)
    Let Balance = Round(Balance / 2, 0)
End Property


Private Property Get Jumping() As String
    Let Jumping = Val(mStrength) + Val(mReactions) + Val(mHeadingAcc)
    Let Jumping = Round(Jumping / 3, 0)
End Property


Private Property Get Vision() As String
    Let Vision = Val(mBallControl) + Val(mReactions)
    Let Vision = Round(Vision / 2, 0)
End Property


Private Property Get SlidingTackle() As String
    Let SlidingTackle = Val(mMarking) + Val(mStandingTackle)
    Let SlidingTackle = Round(SlidingTackle / 2, 0)
End Property


Public Property Get GetDictionary() As Scripting.Dictionary
Variables:
Set oDic = New Scripting.Dictionary
Set cgf = New clsGlobalFunc

AddItems:
    On Error GoTo ErrHandler
    
    With oDic
        Call .Add("crossing", mCrossing)
        Call .Add("finishing", mFinishing)
        Call .Add("heading_accuracy", mHeadingAcc)
        Call .Add("short_passing", mShortPass)
        Call .Add("volleys", Volleys())
        Call .Add("dribbling", mDribbling)
        Call .Add("curve", Curve())
        Call .Add("fk_accuracy", mFkAcc)
        Call .Add("long_passing", mLongPass)
        Call .Add("ball_control", mBallControl)
        Call .Add("acceleration", mAcceleration)
        Call .Add("sprint_speed", mSprintSpeed)
        Call .Add("agility", Agility())
        Call .Add("reactions", mReactions)
        Call .Add("balance", Balance())
        Call .Add("shot_power", mShotPower)
        Call .Add("jumping", Jumping())
        Call .Add("stamina", mStamina)
        Call .Add("strength", mStrength)
        Call .Add("long_shots", mLongShots)
        Call .Add("aggression", mAggression)
        Call .Add("interceptions", mInterceptions)
        Call .Add("positioning", mPositioning)
        Call .Add("vision", Vision())
        Call .Add("penalties", mPenalties)
        Call .Add("composure", mBallControl)
        Call .Add("marking", mMarking)
        Call .Add("standing_tackle", mStandingTackle)
        Call .Add("sliding_tackle", SlidingTackle())
        Call .Add("gk_diving", mGkDiving)
        Call .Add("gk_handling", mGkHandling)
        Call .Add("gk_kicking", mGkKicking)
        Call .Add("gk_positioning", mGkPositioning)
        Call .Add("gk_reflexes", mGkReflexes)
        Call .Add("player_traits", mTraits)
    End With
    
ReturningValue:
    Set GetDictionary = oDic
    
    On Error GoTo 0
    
Finish:
    On Error Resume Next
    Set oDic = Nothing
    Set cgf = Nothing
    Exit Property
    
ErrHandler:
    Call cgf.DisplayErrMessage("SoFifaSkill", "GetDictionary")
    GoTo Finish

End Property


Public Property Let Crossing(ByVal vNewValue As String)
    Let mCrossing = Val(vNewValue)
End Property


Public Property Let Finishing(ByVal vNewValue As String)
    Let mFinishing = Val(vNewValue)
End Property


Public Property Let HeadingAcc(ByVal vNewValue As String)
    Let mHeadingAcc = Val(vNewValue)
End Property


Public Property Let ShortPass(ByVal vNewValue As String)
    Let mShortPass = Val(vNewValue)
End Property


Public Property Let Dribbling(ByVal vNewValue As String)
    Let mDribbling = Val(vNewValue)
End Property


Public Property Let FkAcc(ByVal vNewValue As String)
    Let mFkAcc = Val(vNewValue)
End Property


Public Property Let LongPass(ByVal vNewValue As String)
    Let mLongPass = Val(vNewValue)
End Property


Public Property Let BallControl(ByVal vNewValue As String)
    Let mBallControl = Val(vNewValue)
End Property


Public Property Let Acceleration(ByVal vNewValue As String)
    Let mAcceleration = Val(vNewValue)
End Property


Public Property Let SprintSpeed(ByVal vNewValue As String)
    Let mSprintSpeed = Val(vNewValue)
End Property


Public Property Let Reactions(ByVal vNewValue As String)
    Let mReactions = Val(vNewValue)
End Property


Public Property Let ShotPower(ByVal vNewValue As String)
    Let mShotPower = Val(vNewValue)
End Property


Public Property Let Stamina(ByVal vNewValue As String)
    Let mStamina = Val(vNewValue)
End Property


Public Property Let Strength(ByVal vNewValue As String)
    Let mStrength = Val(vNewValue)
End Property


Public Property Let LongShots(ByVal vNewValue As String)
    Let mLongShots = Val(vNewValue)
End Property


Public Property Let Aggression(ByVal vNewValue As String)
    Let mAggression = Val(vNewValue)
End Property


Public Property Let Interceptions(ByVal vNewValue As String)
    If (vNewValue = vbNullString) Then
        Let mInterceptions = Val(SlidingTackle())
    Else
        Let mInterceptions = Val(vNewValue)
    End If
End Property


Public Property Let Positioning(ByVal vNewValue As String)
    If (vNewValue = vbNullString) Then
        Let mPositioning = Val(mBallControl) + Val(mReactions) + Val(mAggression)
        Let mPositioning = Round(mPositioning / 3, 0)
    Else
        Let mPositioning = Val(vNewValue)
    End If
End Property


Public Property Let Penalties(ByVal vNewValue As String)
    Let mPenalties = Val(vNewValue)
End Property


Public Property Let Marking(ByVal vNewValue As String)
    Let mMarking = Val(vNewValue)
End Property


Public Property Let StandingTackle(ByVal vNewValue As String)
    Let mStandingTackle = Val(vNewValue)
End Property


Public Property Let GkDiving(ByVal vNewValue As String)
    Let mGkDiving = Val(vNewValue)
End Property


Public Property Let GkHandling(ByVal vNewValue As String)
    Let mGkHandling = Val(vNewValue)
End Property


Public Property Let GkKicking(ByVal vNewValue As String)
    Let mGkKicking = Val(vNewValue)
End Property


Public Property Let GkPositioning(ByVal vNewValue As String)
    Let mGkPositioning = Val(vNewValue)
End Property


Public Property Let GkReflexes(ByVal vNewValue As String)
    Let mGkReflexes = Val(vNewValue)
End Property


Public Property Let Traits(ByVal vNewValue As Variant)
    If (vNewValue = vbNullString) Then
        Let mTraits = "[]"
        Exit Property
    End If
    Dim key As Variant
    Let mTraits = Chr(91)
    For key = 1 To UBound(Split(vNewValue, vbCrLf))
        Let mTraits = mTraits & Chr(34) & Trim(Split(vNewValue, vbCrLf)(key)) & Chr(34) & Chr(44)
    Next key
    Let mTraits = Left(mTraits, Len(mTraits) - 1) & Chr(93)
    Set key = Nothing
End Property

