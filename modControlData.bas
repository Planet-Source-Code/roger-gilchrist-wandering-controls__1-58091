Attribute VB_Name = "modControlData"
Option Explicit

'these are a couple of support routines I built to simplify
'reading some control Properties
'You can expand the basic ideas to test(MoveableControl) or extract(GetCaption) any Property
'without having to worry about controls that lack the property
Public Function GetCaption(c As Control) As String

  ' test whether a control has a caption and get it

  On Error Resume Next ' if error generated then returns empty string
  GetCaption = c.Caption
  On Error GoTo 0

End Function

Public Function MoveableControl(c As Control) As Boolean

  ' test whether a control can be moved
  ' ie Timer and CommenDialog have no Height/Width properties so won't be included

  On Error Resume Next ' if error generated then skip line that sets True
  'MoveableControl = c.Height + c.Width
  'Above is simple test but if you get weird and build a 0X0 control it would fail so;
  'code below might fail then skip or succeed and any value including 0 will set True
  MoveableControl = IsNumeric(c.Height + c.Width)
  On Error GoTo 0

End Function

':)Code Fixer V2.8.3 (4/01/2005 9:07:05 AM) 1 + 31 = 32 Lines Thanks Ulli for inspiration and lots of code.
