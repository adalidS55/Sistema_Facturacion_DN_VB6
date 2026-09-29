Attribute VB_Name = "Module2"
''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------
Option Explicit

Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" ( _
ByVal lpPrevWndFunc As Long, _
ByVal hWnd As Long, _
ByVal Msg As Long, _
ByVal Wparam As Long, _
ByVal Lparam As Long) As Long

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" ( _
ByVal hWnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Public Const MK_CONTROL = &H8
Public Const MK_LBUTTON = &H1
Public Const MK_RBUTTON = &H2
Public Const MK_MBUTTON = &H10
Public Const MK_SHIFT = &H4
Private Const GWL_WNDPROC = -4
Private Const WM_MOUSEWHEEL = &H20A

Private procedimientos As New Collection

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal Wparam As Long, ByVal Lparam As Long) As Long

  Dim LocalPrevWndProc As Long
  LocalPrevWndProc = CLng(procedimientos(CStr(Lwnd)))

  Dim MouseKeys As Long
  Dim Rotation As Long
  Dim Xpos As Long
  Dim Ypos As Long

  If Lmsg = WM_MOUSEWHEEL Then
      MouseKeys = Wparam And 65535
      Rotation = Wparam / 65536
      Xpos = Lparam And 65535
      Ypos = Lparam / 65536
      GetForm(Lwnd).MouseWheel MouseKeys, Rotation, Xpos, Ypos
  End If
  WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, Wparam, Lparam)
End Function

Public Sub WheelHook(PassedForm As Form)
  Dim anterior As Long
  anterior = SetWindowLong(PassedForm.hWnd, GWL_WNDPROC, AddressOf WindowProc)
  procedimientos.Add anterior, CStr(PassedForm.hWnd)
End Sub

Public Sub WheelUnHook(PassedForm As Form)
  Dim anterior As Long
  On Error Resume Next
  anterior = CLng(procedimientos(CStr(PassedForm.hWnd)))
  If anterior <> 0 Then
    SetWindowLong PassedForm.hWnd, GWL_WNDPROC, anterior
    procedimientos.Remove CStr(PassedForm.hWnd)
  End If
End Sub

Private Function GetForm(ByVal hWnd As Long) As Form
  For Each GetForm In Forms
    If GetForm.hWnd = hWnd Then Exit Function
  Next GetForm
  Set GetForm = Nothing
End Function

Public Sub FlexGridScroll(ByRef FG As MSFlexGrid, ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim NewValue As Long
  Dim Lstep As Single

  On Error Resume Next

  With FG
      Lstep = .Height / .RowHeight(0)
      Lstep = Int(Lstep)
      
Do While Not (.RowIsVisible(.TopRow + Lstep))
      Lstep = Lstep - 1
    Loop
    
      If Lstep < 10 Then
          Lstep = 10
      End If
      If Rotation > 0 Then
          NewValue = .TopRow - Lstep
          If NewValue < 1 Then
              NewValue = 1
          End If
      Else
          NewValue = .TopRow + Lstep
          If NewValue > .Rows - 1 Then
              NewValue = .Rows - 1
          End If
      End If
  .TopRow = NewValue
  End With
End Sub


''Parte del Scroll Mouse------------------------------------------------------------------------------------------------------------------------------------------------------

