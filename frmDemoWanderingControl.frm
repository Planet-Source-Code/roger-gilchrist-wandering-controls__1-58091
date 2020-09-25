VERSION 5.00
Begin VB.Form frmDemoWanderingControl 
   Caption         =   "Controls Adrift"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMove 
      Caption         =   "Drunk Walk"
      Height          =   495
      Index           =   9
      Left            =   615
      TabIndex        =   10
      Top             =   495
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Down Left"
      Height          =   495
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   990
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Up Right"
      Height          =   495
      Index           =   7
      Left            =   1230
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Down Right"
      Height          =   495
      Index           =   6
      Left            =   1230
      TabIndex        =   6
      Top             =   990
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Up Left"
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Down"
      Height          =   495
      Index           =   4
      Left            =   615
      TabIndex        =   4
      Top             =   990
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Up"
      Height          =   495
      Index           =   3
      Left            =   615
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Left"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Right"
      Height          =   495
      Index           =   1
      Left            =   1230
      TabIndex        =   1
      Top             =   495
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Lock Down"
      Height          =   495
      Index           =   12
      Left            =   3120
      TabIndex        =   19
      ToolTipText     =   "Click this and the Home position will be reset to current position"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "MoveTo"
      Height          =   495
      Index           =   11
      Left            =   2520
      TabIndex        =   18
      ToolTipText     =   "Left-Click and the control will move to the point or hold down Right button and the control will follow the mouse cursor"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Home"
      Height          =   495
      Index           =   10
      Left            =   1920
      TabIndex        =   15
      ToolTipText     =   "Return control to initial position"
      Top             =   0
      Width           =   615
   End
   Begin VB.CheckBox chkBounce 
      Caption         =   "Edge Bounce"
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "Warning if the control is partially off edge of container when bounce is turned on it will get stuck"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   600
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Stop"
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtInput 
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   5760
      Width           =   6015
   End
   Begin VB.ListBox lstControls 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Frame fraContainer 
      Caption         =   "Frame for lblContained"
      Height          =   3735
      Left            =   3600
      TabIndex        =   12
      Top             =   1080
      Width           =   5055
      Begin VB.Label lblContained 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1920
         TabIndex        =   13
         Top             =   1440
         Width           =   270
      End
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Experiment with the buttons above 2. While the label is moving select              a  control in list below"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label lblDemoMoveTo 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmDemoWanderingControl.frx":0000
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   6360
      Width           =   8175
   End
   Begin VB.Label lblInput 
      Caption         =   "Caption for lblContained"
      Height          =   255
      Left            =   1200
      TabIndex        =   16
      Top             =   5520
      Width           =   2175
   End
End
Attribute VB_Name = "frmDemoWanderingControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MvCtrl     As New clsWanderingControl

Private Sub chkBounce_Click()

  MvCtrl.Bounce = chkBounce.Value = vbChecked

End Sub

Private Sub cmdMove_Click(Index As Integer)

  tmrMove.Enabled = Index <> 0
'set initial directions
  MvCtrl.Mode = Index

End Sub

Private Sub Form_DblClick()

  tmrMove.Enabled = True
  MvCtrl.Mode = MoveToTarget

End Sub

Private Sub Form_Load()

  Dim i         As Long
  Dim strCap    As String
  Dim DemoLblNo As Long

  On Error Resume Next
  Set MvCtrl.MovingControl = lblContained
  MvCtrl.CenterOnContainer
  txtInput.Text = "Cute stuff, eh?"
  With Me
    For i = 0 To .Controls.Count - 1
      If MoveableControl(.Controls(i)) Then
        strCap = GetCaption(.Controls(i))
        If LenB(strCap) Then
          strCap = " " & Chr$(34) & strCap & Chr$(34)
        End If
        If .Controls(i).Index < 0 Then
          lstControls.AddItem .Controls(i).Name & strCap
         Else
          lstControls.AddItem .Controls(i).Name & "(" & .Controls(i).Index & ")" & strCap
        End If
'becuase some controls might not be in list you can't use the ListIndex directly
        lstControls.ItemData(lstControls.NewIndex) = i
'but you can identify one input and find it again later
        If .Controls(i).Name = "lblContained" Then
          DemoLblNo = i
        End If
      End If
    Next i
  End With 'Me
  For i = 0 To lstControls.ListCount - 1
    If lstControls.ItemData(i) = DemoLblNo Then
      lstControls.ListIndex = i
      Exit For
    End If
  Next i
  On Error GoTo 0
  MvCtrl.Mode = DrunkWalk
  chkBounce.Value = vbChecked
  tmrMove.Enabled = True

End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  If Button = vbLeftButton Then
    MvCtrl.MoveTo X, Y
  End If

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

  If Button = vbRightButton Then
    MvCtrl.MoveTo X, Y
  End If

End Sub

Private Sub fraContainer_DblClick()

  tmrMove.Enabled = True
  MvCtrl.Mode = MoveToTarget

End Sub

Private Sub fraContainer_MouseDown(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)

  If Button = vbLeftButton Then
    MvCtrl.MoveTo X, Y
  End If

End Sub

Private Sub fraContainer_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)

  If Button = vbRightButton Then
    MvCtrl.MoveTo X, Y
  End If

End Sub

Private Sub lstControls_Click()

'becuase some controls might not be in list you can't use the ListIndex directly

  Set MvCtrl.MovingControl = Me.Controls(lstControls.ItemData(lstControls.ListIndex))

End Sub

Private Sub tmrMove_Timer()

  MvCtrl.Move

End Sub

Private Sub txtInput_Change()

'for fun you can change the text in the label in the frame

  lblContained.Caption = txtInput.Text
'Note the label is autosize so you need to reset the limits of movement
  MvCtrl.SetLimitsOfMovement

End Sub

':)Code Fixer V2.8.3 (4/01/2005 9:07:05 AM) 2 + 131 = 133 Lines Thanks Ulli for inspiration and lots of code.

