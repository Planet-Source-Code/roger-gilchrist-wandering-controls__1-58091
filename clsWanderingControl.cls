VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWanderingControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'based on the idea of Oscar Medina
'Marquees Marquesinas"
'at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58044&lngWId=1
'Converted to a class with extra movements
'DESCRIPTION
'Properties
'MovingControl  -  Set the control to be moved. If class already has a control it is returned to its home position.
'Mode           -  Set the movement to use. See Enum eMove below
'Bounce         -  Boolean bounce off the container edges or scroll over them.
'
'Public Procedures
'CenterOnContainer    - Place the control in the Center of the container
'Move                 - Make one step as dictated by Mode property
'ReturnToStartPostion - Return to original Position
'                      NOTE if Mode = LockDown then StartPosition will be reset to current position
'SetLimitsOfMovement  - Call this if the container resizes (EG call from Form_ReSize)
'MoveTo               - Set a position for DoMoveTo to target
'
'Private Procedures
'LeftOrRight          - logic for left right movement
'UpOrDown             - logic for up down movement
'DoMoveTo             - logic to move to designated position(DoMoveTo)
Public Enum eMove
  Stopped
  Left2Right
  Right2Left
  Bottom2Top
  Top2Bottom
  DiagDownRight
  DiagUpLeft
  DiagUpRight
  DiagDownLeft
  DrunkWalk
  GoHome
  MoveToTarget
  LockDown
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Stopped, Left2Right, Right2Left, Bottom2Top, Top2Bottom, DiagDownRight, DiagUpLeft, DiagUpRight, DiagDownLeft, DrunkWalk
Private GoHome, MoveToTarget, LockDown
#End If
Private m_MoveControl          As Control ' The control
'
Private bBounce                As Boolean ' Restrict to the Container
Private m_Mode                 As eMove ' The movement which sets the following
Private HDir                   As Boolean ' Directional switches
Private VDir                   As Boolean '
'
Private HomeLeft               As Long    ' Original position
Private HomeTop                As Long
'
Private VerticalTopOffet       As Long    ' Container positional data
Private VerticaInnerEdge       As Long
Private VerticalOuterEdge      As Long
Private HorzLeftOffSet         As Long
Private HorzInnerEdge          As Long
Private HorzOuterEdge          As Long
'
Private mTargetLeft            As Long    ' Store MoveTo values
Private mTargetTop             As Long

Public Property Get Bounce() As Boolean

  Bounce = bBounce

End Property

Public Property Let Bounce(ByVal B As Boolean)

  bBounce = B

End Property

Public Sub CenterOnContainer()

'place the control in the centre of its container

  With m_MoveControl
    .Left = (.Container.Width - .Width) / 2
    .Top = (.Container.Height - .Height) / 2
  End With

End Sub

Private Sub DoMoveTo()

  With m_MoveControl
    If mTargetTop <> .Top Then
      .Top = .Top + IIf(mTargetTop > .Top, 10, -10)
      If Abs(mTargetTop - .Top) < 10 Then
        .Top = mTargetTop
      End If
    End If
    If mTargetLeft <> .Left Then
      .Left = .Left + IIf(mTargetLeft > .Left, 10, -10)
      If Abs(mTargetLeft - .Left) < 10 Then
        .Left = mTargetLeft
      End If
    End If
  End With

End Sub

Private Sub LeftOrRight(LR As Boolean, _
                        Optional ByVal bBounce As Boolean = False)

  With m_MoveControl
    If .Left > HorzOuterEdge Then
      .Left = -.Width
     ElseIf .Left < -.Width Then
      .Left = HorzOuterEdge
    End If
    If bBounce Then
      If .Left <= 0 Then
        LR = Not LR
       ElseIf .Left >= HorzInnerEdge Then
        LR = Not LR
      End If
    End If
    .Left = .Left + IIf(LR, 20, -20)
  End With

End Sub

Public Property Get Mode() As eMove

  Mode = m_Mode

End Property

Public Property Let Mode(m As eMove)

'set the Directional Switches

  m_Mode = m
  Select Case m_Mode
   Case Left2Right, Right2Left
    HDir = Mode = Left2Right
   Case Bottom2Top, Top2Bottom
    VDir = Mode = Top2Bottom
   Case DiagDownRight
    HDir = Not Mode = DiagDownRight
    VDir = Not Mode = DiagDownRight
   Case DiagUpLeft
    HDir = Mode = DiagUpLeft
    VDir = Mode = DiagUpLeft
   Case DiagUpRight
    HDir = Mode = DiagUpRight
    VDir = Not Mode = DiagUpRight
   Case DiagDownLeft
    HDir = Not Mode = DiagDownLeft
    VDir = Mode = DiagDownLeft
   Case GoHome
    ReturnToStartPostion
   Case MoveToTarget
    DoMoveTo
   Case LockDown
    HomeLeft = m_MoveControl.Left
    HomeTop = m_MoveControl.Top
  End Select

End Property

Public Sub Move()

  Select Case m_Mode
   Case Stopped
    CenterOnContainer
   Case Left2Right, Right2Left
    LeftOrRight HDir, bBounce
   Case Bottom2Top, Top2Bottom
    UpOrDown VDir, bBounce
   Case DiagDownRight
    LeftOrRight HDir, bBounce
    UpOrDown VDir, bBounce
   Case DiagUpLeft
    LeftOrRight HDir, bBounce
    UpOrDown VDir, bBounce
   Case DiagUpRight
    LeftOrRight HDir, bBounce
    UpOrDown VDir, bBounce
   Case DiagDownLeft
    LeftOrRight HDir, bBounce
    UpOrDown VDir, bBounce
   Case DrunkWalk
    If Rnd > 0.99 Then
      VDir = Not VDir
    End If
    If Rnd > 0.99 Then
      HDir = Not HDir
    End If
    LeftOrRight HDir, bBounce
    UpOrDown VDir, bBounce
   Case MoveToTarget
    DoMoveTo
   Case LockDown
    HomeLeft = m_MoveControl.Left
    HomeTop = m_MoveControl.Top
  End Select

End Sub

Public Sub MoveTo(ByVal NLeft As Variant, _
                  ByVal NTop As Variant)

'use variants becuase it could be Longs or singles (From MouseDown)

  mTargetLeft = NLeft
  mTargetTop = NTop

End Sub

Public Property Get MovingControl() As Control

  Set MovingControl = m_MoveControl
  SetLimitsOfMovement
  CenterOnContainer

End Property

Public Property Set MovingControl(ctrl As Control)

  If Not m_MoveControl Is Nothing Then
    ReturnToStartPostion
  End If
  Set m_MoveControl = ctrl
  SetLimitsOfMovement
  HomeLeft = ctrl.Left
  HomeTop = ctrl.Top

End Property

Public Sub ReturnToStartPostion()

  m_MoveControl.Left = HomeLeft
  m_MoveControl.Top = HomeTop

End Sub

Public Sub SetLimitsOfMovement()

  With m_MoveControl
    VerticalTopOffet = 0
    VerticalOuterEdge = .Container.Height
    VerticaInnerEdge = .Container.Height - .Height
    HorzLeftOffSet = -.Width
    HorzOuterEdge = .Container.Width
    HorzInnerEdge = .Container.Width - .Width
'some containers need extra settings
    If TypeOf .Container Is Frame Then
'Frames top edge
      VerticalTopOffet = IIf(.Container.BorderStyle = 1, 120, 0)
     ElseIf TypeOf .Container Is Form Then
'Forms need to use the ScaleXXXXXX Properties
      VerticalOuterEdge = .Container.ScaleHeight
      VerticaInnerEdge = .Container.ScaleHeight - .Height
      HorzOuterEdge = .Container.ScaleWidth
      HorzInnerEdge = .Container.ScaleWidth - .Width
    End If
  End With

End Sub

Private Sub UpOrDown(UD As Boolean, _
                     Optional ByVal bBounce As Boolean = False)

  With m_MoveControl
    If .Top < VerticalTopOffet Then
      .Top = VerticalOuterEdge
     ElseIf .Top > VerticalOuterEdge Then
      .Top = VerticalTopOffet
    End If
    If bBounce Then
      If .Top < VerticalTopOffet Or .Top > VerticaInnerEdge Then
        UD = Not UD
      End If
    End If
    .Top = .Top + IIf(UD, 10, -10)
  End With

End Sub

':)Code Fixer V2.8.3 (4/01/2005 9:07:04 AM) 61 + 221 = 282 Lines Thanks Ulli for inspiration and lots of code.

