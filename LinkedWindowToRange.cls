VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LinkedWindowToRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


Const S_OK = &H0
Const STATE_SYSTEM_INVISIBLE = &H8000
Const STATE_SYSTEM_OFFSCREEN = &H10000
Const CHILDID_SELF = 0&
Const NAVDIR_NEXT = 5&
Const NAVDIR_PREVIOUS = 6&
Const NAVDIR_FIRSTCHILD = 7&
Const NAVDIR_LASTCHILD = 8&

Private Enum WinEventObjectId
  OBJID_SELF = 0
  OBJID_SYSMENU = -1
  OBJID_TITLEBAR = -2
  OBJID_MENU = -3
  OBJID_CLIENT = -4
  OBJID_VSCROLL = -5
  OBJID_HSCROLL = -6
  OBJID_SIZEGRIP = -7
  OBJID_CARET = -8
  OBJID_CURSOR = -9
  OBJID_ALERT = -10
  OBJID_SOUND = -11
  OBJID_QUERYCLASSNAMEIDX = -12
  OBJID_NATIVEOM = -16
End Enum

Const WS_EX_LAYERED = &H80000, WS_EX_NOACTIVATE = &H8000000, WS_EX_NOINHERITLAYOUT = &H100000, WS_EX_TRANSPARENT = &H20, WS_EX_DLGMODALFRAME = &H1

#If VBA7 = 0 Then
  Private Enum LongPtr:[_]: End Enum
#End If
Private Const PtrNull As LongPtr = 0

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type Rect
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type GUID
    lData1 As Long
    iData2 As Integer
    iData3 As Integer
    aBData4(0 To 7) As Byte
End Type

Private Type RECT6
  XY(3) As POINTAPI
End Type

Private Enum EnumDirective
  EDUnload = -1
  EDHide
  EDWait
  EDShow
End Enum

#If -VBA7 And -Win64 Then
  Private Declare PtrSafe Function FindWindowExW Lib "user32" (ByVal ParenthWnd As LongPtr, ByVal ChildhWnd As LongPtr, ByVal lclassname As LongPtr, ByVal lcaption As LongPtr) As LongPtr
  Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
  Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
#ElseIf VBA7 Then
  Private Declare PtrSafe Function FindWindowExW Lib "user32" (ByVal ParenthWnd As LongPtr, ByVal ChildhWnd As LongPtr, ByVal lclassname As LongPtr, ByVal lcaption As LongPtr) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As Long
  Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
#Else
  Private Declare Function FindWindowExW Lib "user32" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal lclassname As Long, ByVal lcaption As Long) As Long
  Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As LongPtr, ByVal wFlag As Long) As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Private Declare PtrSafe Function SetParent Lib "user32" (ByVal hwndChild As LongPtr, ByVal hwndParent As LongPtr) As Long
Private Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
Private Declare PtrSafe Function moveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As Rect) As Long
Private Declare PtrSafe Function ScreenToClient Lib "user32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByVal lpiid As LongPtr) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As Long, ByVal riid As LongPtr, ppvObject As Object) As Long
Private Declare PtrSafe Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "USER32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetNextWindow Lib "USER32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "USER32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IIDFromString Lib "ole32.dll" (ByVal lpsz As Long, ByVal lpiid As Long) As Long
Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, ByVal riid As Long, ppvObject As Object)
Private Declare Function OffsetRect Lib "user32" (lpRect As Rect, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
#End If


#If VBA7 Then
Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetAncestor Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function getTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal HWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

#If VBA7 Then
Private Declare PtrSafe Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As Any, ByVal hwnd As LongPtr) As LongPtr
#Else
Private Declare Function IUnknown_GetWindow Lib "shlwapi" Alias "#172" (ByVal pIUnk As Any, ByVal hwnd As Long) As Long
#End If

#If -VBA7 And -Win64 Then
  Private Declare PtrSafe Function MonitorFromPoint64 Lib "user32.dll" Alias "MonitorFromPoint" (ByVal pt As LongPtr, ByVal dwFlags As Long) As LongPtr
  Private Declare PtrSafe Function AccessibleObjectFromPoint64 Lib "oleacc" Alias "AccessibleObjectFromPoint" (ByVal lXY As LongPtr, ppacc As Any, pvarChild As Variant) As Long
  Private Declare PtrSafe Function WindowFromPoint64 Lib "user32" Alias "WindowFromPoint" (ByVal POINT As LongPtr) As LongPtr
#ElseIf VBA7 Then
  Private Declare PtrSafe Function MonitorFromPoint32 Lib "user32.dll" Alias "MonitorFromPoint" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As LongPtr
  Private Declare PtrSafe Function AccessibleObjectFromPoint32 Lib "oleacc" Alias "AccessibleObjectFromPoint" (ByVal lX As Long, ByVal lY As Long, ppacc As Office.IAccessible, pvarChild As Variant) As Long
  Private Declare PtrSafe Function WindowFromPoint32 Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
#Else
  Private Declare Function MonitorFromPoint32 Lib "user32.dll" Alias "MonitorFromPoint" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As Long
  Private Declare Function AccessibleObjectFromPoint32 Lib "Oleacc" Alias "AccessibleObjectFromPoint" (ByVal lX As Long, ByVal lY As Long, ppacc As Office.IAccessible, pvarChild As Variant) As Long
  Private Declare Function WindowFromPoint32 Lib "user32" Alias "WindowFromPoint" ( ByVal xPoint As Long, ByVal yPoint As Long) As Long
#End If

#If Win64 Then
Private Type POINTAPI64
  value As LongPtr
End Type
#End If

Private Const ID_ACCESSIBLE As String = "{618736E0-3C3D-11CF-810C-00AA00389B71}"

Private Type MonitorInfo
  cbSize As Long
  rcMonitor As Rect
  rcWork As Rect
  dwFlags As Long
End Type
#If VBA7 Then
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As LongPtr, ByRef lpmi As LongPtr) As Boolean
#Else
Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal hMonitor As Long, ByRef lpmi As Long) As Boolean
#End If
#If Win64 Then
Private Type POINTAPI_AsLongLong
value As LongPtr
End Type
#End If

Private Enum LWA_FLAGS
    LWA_COLORKEY = &H1&
    LWA_ALPHA = &H2&
End Enum

Private Const GWL_EXSTYLE = -20
Private Const AC_SRC_OVER As Byte = 0
Private Const AC_SRC_ALPHA As Byte = 1

Private Enum ULW_FLAGS
    ULW_COLORKEY = &H1&
    ULW_ALPHA = &H2&
    ULW_OPAQUE = &H4&
    ULW_EX_NORESIZE = &H8&
End Enum

Private Type POINT
    X As Long
    Y As Long
End Type


Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10

Const mgSize = 30
Const BlDelay = 0.4
Dim AppVersionNew As Boolean
Dim Aw As Window
Dim pane As Object, Panes
Dim hXLD As LongPtr
Dim hXL7 As LongPtr
Dim hXL7sc As LongPtr
Dim hXL7SV1 As LongPtr
Dim hXL7SH1 As LongPtr
Dim hXL7SV2 As LongPtr
Dim hXL7SH2 As LongPtr
Dim rcV As Rect, rcH As Rect, tRC As Rect, wRC As Rect, rXL7 As Rect
Dim lzHwnd  As LongPtr
Dim RA, TRA As String
Dim X&, Y&, osX&, osY&
Dim w#, h#, VS#
Dim PA As RECT6
Dim RO As Object, rs As Object, RV As Object, RH As Object
Dim r As Object
Dim i As Byte
Dim idx As Byte
Dim c As Byte
Dim t As Byte
Dim SR As Boolean
Dim sc As Boolean
Dim bV As Boolean
Dim bH As Boolean
Dim DRTL As Boolean 'DisplayRightToLeft
Dim DHSB As Boolean 'DisplayHorizontalScrollBar
Dim DVSB As Boolean 'DisplayVerticalScrollBar
Dim DWB As Boolean 'DisplayWorkbookTabs
Dim bWSPL As Boolean
Dim pcxWidth As Long, pcyHeight As Long, pxLeft As Long, pyTop As Long
Dim fWindow As Object
Dim cc%, awfz As Boolean, awspl As Boolean, xx&(1 To 3), Y2&
Dim k%, j%, vi As Boolean
Dim n&, s$, ast As LongPtr, o1
Dim tt!
Dim xSpV As Long, ySpH As Long
Dim oIAc As IAccessible, vKid, WBobj As Object
Dim winWork As Boolean
Dim ir%, mg&
Dim IAcSNB  As IAccessible, SNBTop& ' Sheet Navigation Bar
Dim LL As LongPtr
Dim wLinked As Boolean
Dim nPosition
Const xxxxx = 99999

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub
Property Let linkedWindow(Linked As Boolean)
  wLinked = Linked
  If lzHwnd <> 0 Then
    If GetAncestor(lzHwnd, 3) <> hXL7 Then
      If wLinked Then SetParent lzHwnd, hXL7
    Else
      If Not wLinked Then SetParent lzHwnd, 0
    End If
  End If
  
End Property
Property Set newForm(ByVal Form As Object)
  Set fWindow = Form: formHandle Form
End Property
Property Let newWindow(ByVal Window As LongPtr)
  lzHwnd = Window
End Property

Sub hideTitle()
  showFormTitle False
End Sub

Sub showTitle()
  showFormTitle True
End Sub
Sub Show(ByVal target As Range, RangePosition&)
  nPosition = RangePosition
  GetWindowRect lzHwnd, wRC
  setWindowPositionByRange target
  ShowWindow lzHwnd, 5
End Sub
Sub ReShow()
  SetNewPosition
  ShowWindow lzHwnd, 5
End Sub
Sub Hide()
  ShowWindow lzHwnd, 0
End Sub

Sub Offset(Optional offsetX&, Optional offsetY&)
  osX = offsetX: osY = offsetY
End Sub

Private Sub showFormTitle(Optional hasTitle As Boolean)
  If lzHwnd = 0 Then Exit Sub
  Dim l As LongPtr: l = GetWindowLong(lzHwnd, -16)
  SetWindowLong lzHwnd, -16, IIf(hasTitle, l Or &HC00000, l And Not &HC00000)
  l = GetWindowLong(lzHwnd, -20)
  SetWindowLong lzHwnd, -20, IIf(hasTitle, l Or &H1, l And Not &H1)
  DrawMenuBar lzHwnd
End Sub
Private Sub formHandle(ByVal Form As Object)
  IUnknown_GetWindow Form, VarPtr(lzHwnd)
End Sub
Sub SetTitle(ByVal sTitle$)
  DefWindowProc lzHwnd, &HC, 0, StrPtr(sTitle)
End Sub

Private Sub setWindowPositionByRange(ByVal target As Range, Optional ByVal win As Object)
  hXL7SV1 = 0: hXL7SH1 = 0: hXL7SV2 = 0: hXL7SH2 = 0

  If win Is Nothing Then Set win = ActiveWindow
  Select Case TypeName(win)
  Case "Window": winWork = True
  Case "ProtectedViewWindow":  winWork = False
  Case Else: Exit Sub
  End Select
  Dim p0 As RECT6: PA = p0
  
  Set Aw = win
  AppVersionNew = Val(win.Parent.Parent.Version) > 14
  hXL7 = apiWindowWorksheet(win, True)
  If AppVersionNew Then
    mg = 3:  NUIScrollbar
  Else
    mg = 0: hXL7sc = FindWindowEx(hXL7, 0&, "Scrollbar", vbNullString)
  End If
  GetWindowRect hXL7, rXL7
  GetWindowRect GetAncestor(hXL7, 1), tRC
  With rXL7
    If .Left < tRC.Left Then .Left = tRC.Left
    If .Top < tRC.Top Then .Top = tRC.Top
    AccessibleObjectFromXY .Left + 3, .Top + 3, oIAc, vKid
  End With
  SetNewPosition target
End Sub

Private Sub SetNewPosition(Optional target As Range)
  On Error GoTo e
  Set r = Nothing
  If Not target Is Nothing Then Set RA = target

  With Aw
    Set Panes = .Panes
    Set pane = .Panes(1)
    awfz = .FreezePanes
    awspl = .Split
    idx = .ActivePane.index: SR = .SplitRow: sc = .SplitColumn: c = .Panes.Count
    bWSPL = (SR Or sc) And Not awfz
  End With
  Call setPosition
Exit Sub
e:
  Debug.Print "SetNewPosition: " & Erl & "   " & Err.description
End Sub

Private Function setPosition() As Boolean
  On Error GoTo e
  With pane
    Set rs = .VisibleRange
    X = .PointsToScreenPixelsX(rs.Left)
    Y = .PointsToScreenPixelsY(rs.Top)
  End With
  If bWSPL Then
    With Panes(Panes.Count)
      '' TrýÌ 3 là do ðýõÌng keÒ Split dày 3
      If sc Then xSpV = .PointsToScreenPixelsX(.VisibleRange.Left) - 3
      If SR Then ySpH = .PointsToScreenPixelsY(.VisibleRange.Top) - 3
    End With
  End If
  Select Case TypeName(RA)
  Case "Range": Set r = RA: ir = r.areas.Count: If ir > 1 Then Set r = RA.areas(1)
  Case Else: GoTo h
  End Select
  '//'//'//'//'//'//'//'//'//'//
  XL7HandleInfo
  Call Excel7SelectionPositions
n:

  
  Dim w&, h&
  w = wRC.Right - wRC.Left
  h = wRC.Bottom - wRC.Top

  With PA
    If (nPosition And RPE_FullScreen) = RPE_FullScreen Then
      GoSub lk
      moveWindow lzHwnd, .XY(0).X, .XY(0).Y, .XY(3).X - .XY(0).X, .XY(3).Y - .XY(0).Y, True
    Else
      Dim b As Rect, c As MonitorInfo, hMonitor As LongPtr
      hMonitor = MonitorFromCursor
      c = apiGetMonitorInfo(hMonitor): b = c.rcWork
      If wLinked Then
        GoSub lk
      End If
      If (nPosition And RPE_leftBottom) = RPE_leftBottom Then
        X = .XY(1).X: Y = .XY(2).Y
      ElseIf (nPosition And RPE_RightTop) = RPE_RightTop Then
        X = .XY(2).X: Y = .XY(1).Y
      ElseIf (nPosition And RPE_RightBottom) = RPE_RightBottom Then
        X = .XY(2).X: Y = .XY(2).Y
      Else
        X = .XY(1).X: Y = .XY(1).Y
      End If
      If (nPosition And RPE_WindowLeftAbove) = RPE_WindowLeftAbove Then
        X = X - w: Y = Y - h
      ElseIf (nPosition And RPE_WindowLeftBelow) = RPE_WindowLeftBelow Then
        X = X - w:
      ElseIf (nPosition And RPE_WindowRightAbove) = RPE_WindowRightAbove Then
        Y = Y - h
      Else
      End If
      X = X + osX: Y = Y + osY
      If wLinked Then
        If X + w > .XY(3).X Then X = .XY(3).X - w
        If X < .XY(0).X Then X = .XY(0).X
        If Y + h > .XY(3).Y Then Y = .XY(3).Y - h
        If Y < .XY(0).Y Then Y = .XY(0).Y
        .XY(0).X = X: .XY(0).Y = Y
      Else
        If X + w > b.Right Then X = b.Right - w - 5
        If X < b.Left Then X = b.Left + 5
        If Y + h > b.Bottom Then Y = b.Bottom - h - 5
        If Y < b.Top Then Y = b.Top - 5
        .XY(0).X = X: .XY(0).Y = Y
        GoSub lk
      End If
      SetWindowPos lzHwnd, 0, .XY(0).X, .XY(0).Y, 0, 0, SWP_NOSIZE
    End If
  End With
  
  setPosition = True
l:
Exit Function
h:
  GoTo l
Exit Function
e:
  Debug.Print "SetPosition: " & Err.description
  GoTo l
Exit Function
lk:
  If wLinked Then
    For i = 0 To 3
      ScreenToClient hXL7, PA.XY(i)
    Next
  End If
Return
End Function

Private Function Excel7SelectionPositions() As Integer
On Error GoTo TestErr
' Last Edit: 03/07/2024

      Dim p As RECT6, l As Long: xx(1) = xxxxx: xx(2) = xxxxx: xx(3) = xxxxx: Y2 = xxxxx:
      If AppVersionNew Then
        GetWindowRect hXL7SV1, rcV
        GetWindowRect hXL7SH1, rcH
        If hXL7SV2 > 0 Then GetWindowRect hXL7SV2, tRC: If tRC.Bottom > rcV.Bottom And tRC.Bottom <= rXL7.Bottom Then rcV = tRC
        If hXL7SH2 > 0 Then GetWindowRect hXL7SH2, tRC: If tRC.Bottom > rcH.Bottom And tRC.Bottom <= rXL7.Bottom Then rcH = tRC
      Else
        GetWindowRect hXL7sc, rcV
        rcH.Top = rcV.Top + 1
        rcH.Right = rcV.Left + 1
        rcV.Bottom = rcV.Top + 1
      End If
      If bWSPL Then cc = idx: c = idx Else cc = 1: idx = c
      '//'//'//'//'//'//'//'//'//'//'//'//

      With p
        If DRTL Then
          If DVSB Then .XY(0).X = rcV.Right + mg Else .XY(0).X = rXL7.Left
          .XY(3).X = rXL7.Right - (X - rXL7.Left)
        Else
          .XY(0).X = X
          If DVSB Then .XY(3).X = rcV.Left - mg Else .XY(3).X = rXL7.Right + 4
        End If
        .XY(0).Y = Y
        Select Case True
        Case DWB: .XY(3).Y = SNBTop
        Case DVSB: .XY(3).Y = rcV.Bottom - 2
        Case DHSB: .XY(3).Y = rcH.Top - mg - 1
        Case Else: .XY(3).Y = rXL7.Bottom
        End Select
      End With
      '//'//'//'//'//'//'//'//'//'//'//'//
      t = 0: bV = False: bH = False: w = 0: h = 0
      For i = cc To c
        With Panes(i)
          Set rs = .VisibleRange
          With .Application

            If (Not awfz) And awspl And idx <> i Then GoTo Next1
s:
            Set RO = .Intersect(r, rs):
            Set RV = .Intersect(r.EntireColumn, rs)
            Set RH = .Intersect(r.EntireRow, rs)
          End With
          If Not rs Is Nothing Then
            If bWSPL Then
              Select Case i
              Case 1, 3:
                xx(2 + DRTL) = .PointsToScreenPixelsX(rs(1, rs.Columns.Count + 1).Left):
                If sc And bWSPL Then
                  Select Case i
                  Case 1, 3:
                    If xx(2 + DRTL) > xSpV Then xx(2 + DRTL) = xSpV:
                    If p.XY(2).X > xSpV Then p.XY(2).X = xSpV:
                  End Select
                End If
                Y2 = .PointsToScreenPixelsY(rs(rs.Rows.Count + 1, 1).Top):
                If SR And bWSPL Then
                  Select Case i
                  Case 1:
                    If Y2 > ySpH Then Y2 = ySpH
                    If p.XY(2).Y > ySpH Then p.XY(2).Y = ySpH:
                  End Select
                End If
              End Select
            End If
          End If
          If Not RV Is Nothing Then
            If Not bV Then bV = True: p.XY(1 - DRTL).X = .PointsToScreenPixelsX(RV.Left)
            If t = 0 Or (sc And t = 1 And i = 2) Or (sc And t = 3 And i = 4) Then
              p.XY(2 + DRTL).X = .PointsToScreenPixelsX(RV(1, RV.Columns.Count + 1).Left)
            End If
          End If
          If Not RH Is Nothing Then
            If Not bH Then bH = True: p.XY(1).Y = .PointsToScreenPixelsY(RH.Top)
            If t = 0 _
            Or (SR And c = 2 And t = 1 And i = 2) _
            Or (c = 4 And t = 1 And i = 3) _
            Or (c = 4 And t = 2 And i = 4) Then
              p.XY(2).Y = .PointsToScreenPixelsY(RH(RH.Rows.Count + 1, 1).Top)
            End If
          End If
        End With
        If t = 0 Then t = IIf(Not RO Is Nothing, i, 0)
    
Next1:
  Next i

  With p
    If Y2 <> xxxxx Then If Y2 < .XY(2).Y Then .XY(2).Y = Y2
    If DRTL Then
      ' ÐaÒo viò trí
      .XY(2).X = .XY(3).X - .XY(2).X + X + 1
      .XY(1).X = .XY(3).X - .XY(1).X + X + 1
      If xx(1) <> xxxxx Then xx(1) = .XY(3).X - xx(1) + X + 1: If xx(1) > .XY(1).X Then .XY(1).X = xx(1) - 2
      If .XY(1).X < .XY(0).X Then .XY(1).X = .XY(0).X
    Else
      If xx(2) <> xxxxx Then If xx(2) < .XY(2).X Then .XY(2).X = xx(2) + 3
    End If
    If .XY(2).X > .XY(3).X Then .XY(2).X = .XY(3).X
    If .XY(2).Y > .XY(3).Y Then .XY(2).Y = .XY(3).Y
  End With
  PA = p
  Excel7SelectionPositions = 1
e:
Exit Function:
TestErr: Excel7SelectionPositions = -1
  #If Devcoding Then
    Debug.Print "Positions: "; i; Erl, rs.Address, Err.description
  #End If
End Function
Sub HideNow()
  If IsWindow(lzHwnd) <> 0 Then ShowWindow lzHwnd, 0
End Sub
Sub terminate()
  On Error Resume Next
  If IsWindow(lzHwnd) <> 0 Then Call HideNow: SetParent lzHwnd, 0
  Set oIAc = Nothing
End Sub
Private Sub NUIScrollbar()
  If hXL7SV2 <> 0 And hXL7SH2 <> 0 Then Exit Sub
  Dim n As LongPtr, v$, l&
  n = FindWindowEx(hXL7, 0&, vbNullString, vbNullString)
  hXL7SV1 = 0: hXL7SH1 = 0: hXL7SV2 = 0: hXL7SH2 = 0
  Do Until n = 0
    v = Space$(100): l = GetWindowText(n, v, 100): If l Then v = Left$(v, l) Else s = vbNullString
    Select Case True
    Case v = "Vertical", v Like "D?c": If hXL7SV1 = 0 Then hXL7SV1 = n Else hXL7SV2 = n
    Case v = "Horizontal", v = "Ngang": If hXL7SH1 = 0 Then hXL7SH1 = n Else hXL7SH2 = n
    End Select
    n = FindWindowEx(hXL7, n, vbNullString, vbNullString)
  Loop
End Sub
Private Sub XL7HandleInfo()
  On Error Resume Next
  NUIScrollbar
  Dim i&, Y As Boolean, xc%, o2
  With Aw
    If winWork Then
      DRTL = .DisplayRightToLeft
      DVSB = .DisplayVerticalScrollBar
      DHSB = .DisplayHorizontalScrollBar
      'DWB = .DisplayWorkbookTabs
    Else
      DRTL = .Workbook.ActiveSheet.DisplayRightToLeft
      DVSB = IsWindowVisible(hXL7SV1) <> 0 Or IsWindowVisible(hXL7SV2) <> 0
      DHSB = IsWindowVisible(hXL7SH1) <> 0 Or IsWindowVisible(hXL7SH2) <> 0
    End If
  End With
  If IAcSNB Is Nothing Then
    GoSub byPos
  Else
    GoSub SNB:
  End If
n:
Exit Sub
SNB:
  Set o1 = IAcSNB:
  ast = o1.accState(CHILDID_SELF)
  DWB = ((ast And STATE_SYSTEM_INVISIBLE) <> STATE_SYSTEM_INVISIBLE) _
            And ((ast And STATE_SYSTEM_OFFSCREEN) <> STATE_SYSTEM_OFFSCREEN)
  If DWB Then GoSub lc:  SNBTop = pyTop
Return
byPos:
  Set o1 = oIAc:  GoSub nac
Return
nac: DWB = False: n = NAVDIR_FIRSTCHILD

lnac:
  Set o1 = oIAc
  i = o1.accParent.accChildCount:
  For k = 1 To i
    Err.Clear:  Set o1 = o1.accNavigate(n, CHILDID_SELF)
    If Err Then
      Exit For
    Else
      If n = NAVDIR_FIRSTCHILD Then
        n = NAVDIR_NEXT
      ElseIf n = NAVDIR_LASTCHILD Then
        n = NAVDIR_PREVIOUS
      End If
    End If
    s = o1.accName(CHILDID_SELF): ast = o1.accState(CHILDID_SELF)
    If s <> Empty Then
      vi = ((ast And STATE_SYSTEM_INVISIBLE) <> STATE_SYSTEM_INVISIBLE) _
            And ((ast And STATE_SYSTEM_OFFSCREEN) <> STATE_SYSTEM_OFFSCREEN)
      If vi Then
        Select Case True
        Case s = "Sheet Navigation Bar", s Like "Thanh D??n h???ng Trang t??nh":  GoSub lc:
          If (pyTop >= rXL7.Bottom - 50) And (pyTop + pcyHeight >= rXL7.Bottom - 50) Then
            DWB = True: Set IAcSNB = o1: SNBTop = pyTop
          End If
        End Select
      End If
    End If
  Next
  If DWB = False And n = NAVDIR_NEXT Then: n = NAVDIR_LASTCHILD: GoTo lnac
Return
nlc: n = NAVDIR_LASTCHILD: GoSub a: Return
nfc: n = NAVDIR_FIRSTCHILD: GoSub a: Return
nn: n = NAVDIR_NEXT: GoSub a: Return
np: n = NAVDIR_PREVIOUS: GoSub a: Return
a:
  Set o1 = o1.accNavigate(n, CHILDID_SELF)
Return
lc:
  Call o1.accLocation(pxLeft, pyTop, pcxWidth, pcyHeight, CHILDID_SELF)
Return
End Sub
Private Function getAccLocation(Acc As IAccessible, Optional visible As Boolean, Optional hwndScreenToClient As LongPtr) As Rect
  Dim l&, t&, w&, h&
  t = Acc.accState(CHILDID_SELF)
  If ((t And STATE_SYSTEM_INVISIBLE) = STATE_SYSTEM_INVISIBLE) _
            Or ((t And STATE_SYSTEM_OFFSCREEN) = STATE_SYSTEM_OFFSCREEN) Then Exit Function
  visible = True
  Call Acc.accLocation(l&, t&, w&, h&, CHILDID_SELF)
  If hwndScreenToClient > 0 Then
    Dim p As POINTAPI: p.X = l: p.Y = t
    ScreenToClient hwndScreenToClient, p: l = p.X: t = p.Y
  End If
  With getAccLocation
    .Left = l: .Top = t: .Right = l + w: .Bottom = t + h
  End With
End Function

Sub Transparent(percent!)
  percent = IIf(percent < 0.05, 0.05, IIf(percent > 1, 1, percent))
  setTransparent percent * 255
End Sub
Private Sub setTransparent(Optional ByVal opacity As Byte = 255)
  SetWindowLong lzHwnd, -20, GetWindowLong(lzHwnd, -20) Or WS_EX_LAYERED
  SetLayeredWindowAttributes lzHwnd, 0&, opacity, 2
End Sub
Public Function WindowFromXY(ByVal X As Long, ByVal Y As Long) As LongPtr
  Dim p As POINTAPI: p.X = X: p.Y = Y: WindowFromXY = WindowFromPoint(p)
End Function
Friend Property Get WindowFromPoint(pt As POINTAPI) As LongPtr
  #If Win64 Then
    Dim t As POINTAPI64: LSet t = pt
    WindowFromPoint = WindowFromPoint64(t.value)
  #Else
    WindowFromPoint = WindowFromPoint32(pt.X, pt.Y)
  #End If
End Property
Public Function AccessibleObjectFromXY(ByVal X As Long, ByVal Y As Long, ppac As Object, varchild As Variant) As Long
  Dim p As POINTAPI: p.X = X: p.Y = Y: AccessibleObjectFromXY = AccessibleObjectFromPoint(p, ppac, varchild)
End Function
Friend Function AccessibleObjectFromPoint(pt As POINTAPI, ppac As Object, varchild As Variant) As Long
  #If Win64 Then
    Dim t As POINTAPI64: LSet t = pt
    AccessibleObjectFromPoint = AccessibleObjectFromPoint64(t.value, ppac, varchild)
  #Else
    AccessibleObjectFromPoint = AccessibleObjectFromPoint32(pt.X, pt.Y, ppac, varchild)
  #End If
End Function
Private Function apiWindowWorksheet(ByVal win As Window, Optional xl7 As Boolean) As LongPtr
  Dim h As LongPtr:
  Const n_ = vbNullString
  If AppVersionNew Then
    h = win.hwnd: If xl7 Then h = FindWindowEx(FindWindowEx(h, 0&, "XLDESK", n_), 0&, "EXCEL7", n_)
  Else
    Dim z As LongPtr, l As LongPtr, s$:
    s = win.caption: l = StrPtr("EXCEL7")
    z = FindWindowEx(win.ActiveSheet.Parent.Parent.hwnd, 0&, "XLDESK", n_)
    h = FindWindowExW(z, 0&, l, StrPtr(s))
    If h = 0 Then
      h = FindWindowExW(z, 0&, l, StrPtr(s & "  [Read-Only]")):
      If h = 0 Then
        h = FindWindowExW(z, 0&, l, StrPtr(s & "  [Repair]")):
        If h = 0 Then h = FindWindowExW(z, 0&, l, StrPtr(s & "  [Repaired]"))
      End If
    End If
    If h = 0 Then h = FindWindowExW(z, 0&, l, n_)
  End If
  apiWindowWorksheet = h
End Function
Private Sub UserForm_Terminate()
  Set Aw = Nothing
  Set RO = Nothing
  Set rs = Nothing
  Set RV = Nothing
  Set RH = Nothing
  Set fWindow = Nothing
  Set IAcSNB = Nothing
End Sub
Private Function apiGetMonitorInfo(Optional ByVal monitor As LongPtr) As MonitorInfo
  If monitor = 0 Then monitor = MonitorFromCursor
  apiGetMonitorInfo.cbSize = Len(apiGetMonitorInfo)
  Call GetMonitorInfo(monitor, ByVal VarPtr(apiGetMonitorInfo))
End Function
Private Function MonitorFromCursor() As LongPtr
  Dim p As POINTAPI: GetCursorPos p: MonitorFromCursor = MonitorFromPoint(p, 0)
End Function
#If Win64 Then
Friend Function MonitorFromPoint(pt As POINTAPI, Optional ByVal dwFlags As Long) As LongPtr
  Dim t As POINTAPI_AsLongLong: LSet t = pt
  MonitorFromPoint = MonitorFromPoint64(t.value, dwFlags)
End Function
#Else
Friend Function MonitorFromPoint(pt As POINTAPI, Optional ByVal dwFlags As Long) As LongPtr
  MonitorFromPoint = MonitorFromPoint32(pt.X, pt.Y, dwFlags)
End Function
#End If
