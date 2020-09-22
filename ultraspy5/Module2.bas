Attribute VB_Name = "mdlGetStyles"
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)


Public Type STYLEBUF
        dwStyle As Long
        szDescription As String
End Type
  
Public Const WS_BORDER = &H800000
Public Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Public Const WS_CHILD = &H40000000
Public Const WS_CHILDWINDOW = (WS_CHILD)
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_DLGFRAME = &H400000
Public Const WS_GROUP = &H20000
Public Const WS_HSCROLL = &H100000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_ICONIC = WS_MINIMIZE
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_OVERLAPPED = &H0&
Public Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Public Const WS_POPUP = &H80000000
Public Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Public Const WS_SIZEBOX = WS_THICKFRAME
Public Const WS_TABSTOP = &H10000
Public Const WS_TILED = WS_OVERLAPPED
Public Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Public Const WS_VISIBLE = &H10000000
Public Const WS_VSCROLL = &H200000

Public Const WS_EX_ACCEPTFILES = &H10&
Public Const WS_EX_DLGMODALFRAME = &H1&
Public Const WS_EX_NOPARENTNOTIFY = &H4&
Public Const WS_EX_TOPMOST = &H8&
Public Const WS_EX_TRANSPARENT = &H20&





Public Sub GetStyles(hWnd1 As Long, retString As String)
Dim wStyle As Long

wStyle = GetWindowLong(hWnd1, GWL_STYLE)


'one by one we step through wStyle to see if it has the constant in it
'if the style exists in wStyle, it is added to a string which is returned
If (wStyle And WS_BORDER) Then retString = retString & "WS_BORDER" & " - "
If (wStyle And WS_CAPTION) Then retString = retString & "WS_CAPTION" & " - "
If (wStyle And WS_CHILD) Then retString = retString & "WS_CHILD" & " - "
If (wStyle And WS_CHILDWINDOW) Then retString = retString & "WS_CHILDWINDOW" & " - "
If (wStyle And WS_CLIPCHILDREN) Then retString = retString & "WS_CLIPCHILDREN" & " - "
If (wStyle And WS_CLIPSIBLINGS) Then retString = retString & "WS_CLIPSIBLINGS" & " - "
If (wStyle And WS_DISABLED) Then retString = retString & "WS_DISABLED" & " - "
If (wStyle And WS_DLGFRAME) Then retString = retString & "WS_DLGFRAME" & " - "
If (wStyle And WS_GROUP) Then retString = retString & "WS_GROUP" & " - "
If (wStyle And WS_HSCROLL) Then retString = retString & "WS_HSCROLL" & " - "
If (wStyle And WS_MINIMIZE) Then retString = retString & "WS_MINIMIZE" & " - "
If (wStyle And WS_ICONIC) Then retString = retString & "WS_ICONIC" & " - "
If (wStyle And WS_MAXIMIZE) Then retString = retString & "WS_MAXIMIZE" & " - "
If (wStyle And WS_MAXIMIZEBOX) Then retString = retString & "WS_MAXIMIZEBOX" & " - "
If (wStyle And WS_SYSMENU) Then retString = retString & "WS_SYSMENU" & " - "
If (wStyle And WS_THICKFRAME) Then retString = retString & "WS_THICKFRAME" & " - "
If (wStyle And WS_MINIMIZEBOX) Then retString = retString & "WS_MINIMIZEBOX" & " - "
If (wStyle And WS_OVERLAPPED) Then retString = retString & "WS_OVERLAPPED" & " - "
If (wStyle And WS_OVERLAPPEDWINDOW) Then retString = retString & "WS_OVERLAPPEDWINDOW" & " - "
If (wStyle And WS_POPUP) Then retString = retString & "WS_POPUP" & " - "
If (wStyle And WS_POPUPWINDOW) Then retString = retString & "WS_POPUPWINDOW" & " - "
If (wStyle And WS_SIZEBOX) Then retString = retString & "WS_SIZEBOX" & " - "
If (wStyle And WS_TABSTOP) Then retString = retString & "WS_TABSTOP" & " - "
If (wStyle And WS_TILED) Then retString = retString & "WS_TILED" & " - "
If (wStyle And WS_TILEDWINDOW) Then retString = retString & "WS_TILEDWINDOW" & " - "
If (wStyle And WS_VISIBLE) Then retString = retString & "WS_VISIBLE" & " - "
If (wStyle And WS_VSCROLL) Then retString = retString & "WS_VSCROLL" & " - "


End Sub


