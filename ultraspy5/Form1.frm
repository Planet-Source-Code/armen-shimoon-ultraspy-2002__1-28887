VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3915
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Form1.frx":212A
   ScaleHeight     =   3915
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picShow 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1800
      Picture         =   "Form1.frx":29F4
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   44
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picGen 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1530
      Picture         =   "Form1.frx":2AE0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   43
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2880
      Top             =   3420
   End
   Begin VB.PictureBox picHelp 
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   1260
      Picture         =   "Form1.frx":2BCC
      ScaleHeight     =   173.333
      ScaleMode       =   0  'User
      ScaleWidth      =   162.5
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picSave 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   990
      Negotiate       =   -1  'True
      Picture         =   "Form1.frx":2D20
      ScaleHeight     =   160
      ScaleMode       =   0  'User
      ScaleWidth      =   162.5
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Pointer"
      Height          =   1545
      Left            =   3240
      TabIndex        =   21
      Top             =   2250
      Width           =   2985
      Begin VB.TextBox Label6 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   450
         Width           =   645
      End
      Begin VB.TextBox Label12 
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   450
         Width           =   1365
      End
      Begin VB.TextBox txtPos 
         Height          =   285
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1080
         Width           =   2445
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "RGB"
         Height          =   285
         Left            =   1530
         TabIndex        =   29
         Top             =   180
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Color:"
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   450
         Width           =   465
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Cursor Position"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   810
         Width           =   2355
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Edit Window"
      Height          =   1995
      Left            =   3240
      TabIndex        =   11
      Top             =   180
      Width           =   2985
      Begin VB.CommandButton cmdNotOnTop 
         Caption         =   "Not Ontop"
         Height          =   465
         Left            =   2070
         TabIndex        =   38
         Top             =   1350
         Width           =   645
      End
      Begin VB.CommandButton cmdOnTop 
         Caption         =   "Ontop"
         Height          =   465
         Left            =   1440
         TabIndex        =   37
         Top             =   1350
         Width           =   645
      End
      Begin VB.CommandButton cmdFocus 
         Caption         =   "Focus"
         Height          =   465
         Left            =   2070
         TabIndex        =   36
         Top             =   900
         Width           =   645
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2250
         TabIndex        =   35
         Top             =   180
         Width           =   285
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Height          =   465
         Left            =   810
         TabIndex        =   34
         Top             =   1350
         Width           =   645
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   465
         Left            =   180
         TabIndex        =   33
         Top             =   1350
         Width           =   645
      End
      Begin VB.CommandButton cmdMax 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1980
         TabIndex        =   32
         Top             =   180
         Width           =   285
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   31
         Top             =   180
         Width           =   285
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "Flash"
         Height          =   465
         Left            =   1440
         TabIndex        =   30
         Top             =   900
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   180
         TabIndex        =   15
         Top             =   540
         Width           =   2625
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Enable"
         Height          =   465
         Left            =   810
         TabIndex        =   14
         Top             =   900
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Disable"
         Height          =   465
         Left            =   180
         TabIndex        =   13
         Top             =   900
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1710
         TabIndex        =   12
         Top             =   180
         Width           =   285
      End
      Begin VB.Label Label7 
         Caption         =   "New Text:"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Window Information"
      Height          =   3615
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2985
      Begin VB.TextBox txtHandle 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1815
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   590
         Left            =   2160
         ScaleHeight     =   525
         ScaleWidth      =   525
         TabIndex        =   41
         Top             =   2880
         Width           =   590
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   900
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   25
         Top             =   2880
         Width           =   555
      End
      Begin VB.TextBox txtPClass 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2070
         Width           =   1815
      End
      Begin VB.TextBox txtPID 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2430
         Width           =   1815
      End
      Begin VB.TextBox txtParent 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1350
         Width           =   1815
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   1815
      End
      Begin VB.TextBox txtUM 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   990
         Width           =   1815
      End
      Begin VB.TextBox txtPText 
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1710
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "Icon:"
         Height          =   285
         Left            =   1710
         TabIndex        =   40
         Top             =   2970
         Width           =   465
      End
      Begin VB.Label Label13 
         Caption         =   "Drag:"
         Height          =   285
         Left            =   180
         TabIndex        =   26
         Top             =   2970
         Width           =   645
      End
      Begin VB.Label Label10 
         Caption         =   "PClass:"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label9 
         Caption         =   "Process ID:"
         Height          =   375
         Left            =   180
         TabIndex        =   17
         Top             =   2430
         Width           =   555
      End
      Begin VB.Label Label11 
         Caption         =   "Handle:"
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "PHandle:"
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Text:"
         Height          =   180
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Class:"
         Height          =   270
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Label15 
         Caption         =   "PText:"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   1800
         Width           =   555
      End
   End
   Begin VB.Label Label16 
      Caption         =   "Build 01"
      Height          =   255
      Left            =   5640
      TabIndex        =   45
      Top             =   0
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   6660
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   6570
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveInfo 
         Caption         =   "Save Window &Information"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuIcon 
         Caption         =   "Save &Window Icon"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuDC 
         Caption         =   "Show &DC Info"
      End
      Begin VB.Menu mnuStyles 
         Caption         =   "Show &Styles Info"
      End
      Begin VB.Menu coolsep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGenSC 
         Caption         =   "&Generate Subclassing Code"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSmall 
         Caption         =   "&Small Mode"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAOT 
         Caption         =   "Always On &Top"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Shimoon Update"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuReadMe 
         Caption         =   "&Readme"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim verylong As String * 100
Dim gParent As String * 100
Dim SndMsg As String * 100
Dim windowname As String * 100
Dim sztext As String * 100
Dim mousemove As Boolean
Dim ParentWindow As String * 255
Dim wndInfo As String
Dim i1 As Integer
Dim WndRECT As RECT

Private Sub cmdAbout_Click()
End Sub



Private Sub cmdMakeSmall_Click()

End Sub

Private Sub cmdClose_Click()
If txtHandle.Text <> "" Then
Call SendMessage(txtHandle.Text, WM_CLOSE, 0, 0)
End If
End Sub

Private Sub cmdFlash_Click()
If txtHandle.Text <> "" Then
Timer1.Enabled = True
i1 = 0
End If
End Sub

Private Sub cmdFocus_Click()
If txtHandle.Text <> "" Then
SetActiveWindow txtHandle.Text
End If
End Sub

Private Sub cmdHide_Click()
If txtHandle.Text <> "" Then
Call ShowWindow(txtHandle.Text, SW_HIDE)
End If
End Sub

Private Sub cmdMax_Click()
If txtHandle.Text <> "" Then
On Error GoTo 50
Call ShowWindow(txtHandle.Text, SW_MAXIMIZE)
End If
50: Exit Sub
End Sub

Private Sub cmdNotOnTop_Click()
If txtHandle.Text <> "" Then
 SetWindowPos txtHandle.Text, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOSIZE
End If
 
End Sub

Private Sub cmdOnTop_Click()
If txtHandle.Text <> "" Then
 SetWindowPos txtHandle.Text, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOSIZE
End If
End Sub

Private Sub cmdRestore_Click()
If txtHandle.Text <> "" Then
Call ShowWindow(txtHandle.Text, SW_RESTORE)
End If
End Sub

Private Sub cmdShow_Click()
If txtHandle.Text <> "" Then
Call ShowWindow(txtHandle.Text, SW_SHOW)
End If
End Sub

Private Sub Command1_Click()
If txtHandle.Text <> "" Then
Call EnableWindow(txtHandle.Text, 1)
End If
End Sub

Private Sub Command2_Click()
If txtHandle.Text <> "" Then
Call EnableWindow(txtHandle.Text, 0)
End If
End Sub

Private Sub Command3_Click()
If txtHandle.Text <> "" Then
Call CloseWindow(txtHandle.Text)
End If
End Sub




Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
mousemove = False ' makes sure program knows we arent dragging the pointer
UVersion = "2002.0"

Dim OnTop As Boolean

If GetSetting("UltraSpy5", "Position", "Top") = "" Then
    Exit Sub
Else
    Form1.Top = GetSetting("Ultraspy5", "Position", "Top")
    Form1.Left = GetSetting("Ultraspy5", "Position", "Left")
    OnTop = GetSetting("Ultraspy5", "Settings", "OnTop")
    AlwaysOnTop Form1, OnTop
    If OnTop = True Then
        Form1.mnuAOT.Checked = True
    End If
    Form1.Width = GetSetting("Ultraspy5", "Settings", "Width")
    If Form1.Width = 3255 Then
        Form1.mnuSmall.Checked = True
    End If
End If


CButton Form1
CButton frmDC

CButton frmSplash
CButton frmStyle
Call DoMenu



    Set fD = New cFileDialog
    fD.Filter = "Text Files (*.txt)|*.txt|Bitmap Files (*.bmp)|*.bmp|Icon Files (*.ico)|*.ico"

With Picture1
    .Width = .Width - 30
    .Height = .Height - 30
End With

txtHandle.SetFocus
SendKeys vbNullString


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim OnTop As Boolean
Dim fWidth As String
Set fD = Nothing

If mnuAOT.Checked = True Then
    OnTop = True
Else
    OnTop = False
End If
    
If mnuSmall.Checked = True Then
    fWidth = "3255"
Else
    fWidth = "6405"
End If

SaveSetting "UltraSpy5", "Position", "Top", Form1.Top
SaveSetting "UltraSpy5", "Position", "Left", Form1.Left
SaveSetting "UltraSpy5", "Settings", "OnTop", OnTop
SaveSetting "UltraSpy5", "Settings", "Width", fWidth


'unloads all the forms from memory
Unload frmSplash
Unload Form1
Unload frmReadme
Unload frmDC
Unload frmStyle
Unload frmUpdate
End
End Sub



Private Sub mnuAbout_Click()
'displays messagbox about this program

Call ShellAbout(Form1.hwnd, "Shimoon UltraSpy 2002", "Made by Armen Shimoon of Shimoon Technologies. © 2000 - 2002.", Form1.Icon)

End Sub

Private Sub mnuAOT_Click()
If mnuAOT.Checked = True Then
    mnuAOT.Checked = False
    AlwaysOnTop Form1, False
Else
    mnuAOT.Checked = True
    AlwaysOnTop Form1, True
End If
End Sub

Private Sub mnuDC_Click()
frmDC.Show
End Sub

Private Sub mnuExit_Click()
Dim OnTop As Boolean
Dim fWidth As String
Set fD = Nothing

If mnuAOT.Checked = True Then
    OnTop = True
Else
    OnTop = False
End If
    
If mnuSmall.Checked = True Then
    fWidth = "3255"
Else
    fWidth = "6240"
End If

SaveSetting "UltraSpy5", "Position", "Top", Form1.Top
SaveSetting "UltraSpy5", "Position", "Left", Form1.Left
SaveSetting "UltraSpy5", "Settings", "OnTop", OnTop
SaveSetting "UltraSpy5", "Settings", "Width", fWidth


'unloads all the forms from memory
Unload frmSplash
Unload Form1
Unload frmReadme
Unload frmDC
Unload frmStyle
Unload frmUpdate
End
End Sub



Private Sub mnuGenSC_Click()
frmSC.Show
End Sub

Private Sub mnuGreetings_Click()
frmAbout.Show
End Sub

Private Sub mnuIcon_Click()

On Error GoTo errorhandler
    With fD
        .flags = OFN_OVERWRITEPROMPT

        .hwnd = Me.hwnd
        .DefaultExt = "*.ico"
        .CancelError = False
        .ShowSave

    End With



SavePicture picIcon.Image, fD.Filename
MsgBox "Icon saved to " & fD.Filename, vbInformation, "Saved"

fD.Filename = ""

Exit Sub
errorhandler:
    MsgBox Err.Description
End Sub

Private Sub mnuReadMe_Click()
frmReadme.Show
End Sub

Private Sub mnuSaveInfo_Click()
wndInfo = wndInfo & "Handle: " & txtHandle.Text & vbCrLf
wndInfo = wndInfo & "Class: " & txtClass.Text & vbCrLf
wndInfo = wndInfo & "Text: " & txtUM.Text & vbCrLf
wndInfo = wndInfo & "Parent Handle: " & txtParent.Text & vbCrLf
wndInfo = wndInfo & "Parent Text: " & txtPText.Text & vbCrLf
wndInfo = wndInfo & "Parent Class: " & txtPClass.Text & vbCrLf
wndInfo = wndInfo & "Process ID: " & txtPID.Text & vbCrLf
wndInfo = wndInfo & "WS_STYLES: " & frmStyle.lblStyle.Text & vbCrLf
wndInfo = wndInfo & "RECT: " & frmStyle.txtRect.Text & vbCrLf



On Error GoTo errorhandler
    With fD
        .flags = OFN_OVERWRITEPROMPT

        .hwnd = Me.hwnd
        .DefaultExt = "txt"
        .CancelError = False
        .ShowSave

    End With
    

Open fD.Filename For Output As #1
    Print #1, wndInfo
Close #1
MsgBox "File saved to " & fD.Filename, vbInformation, "Saved"
fD.Filename = ""
Exit Sub
errorhandler:
    MsgBox Err.Description

End Sub

Private Sub mnuSmall_Click()
If mnuSmall.Checked = True Then
    mnuSmall.Checked = False
    Form1.Width = 6405
Else
    mnuSmall.Checked = True
    Form1.Width = 3255
End If
End Sub

Private Sub mnuStyles_Click()
frmStyle.Show
End Sub

Private Sub mnuUpdate_Click()
frmUpdate.Show
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'gives the effect that we are actually dragging the pointer
Picture1.Picture = Nothing
Form1.MouseIcon = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 99
'tells program to get information
mousemove = True

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim cursorPos1 As POINTAPI
   Dim wintext As String
   Dim garmon As String
   Dim gIcon As Image
   Dim OldX As Integer
   Dim OldY As Integer
   Dim ttxt As String
   Dim abc As String
   Dim h2Icon As Long
   Dim Width1 As Integer, Height1 As Integer
   Dim sWnd As Long
   Dim wStyles As String, wStyles2 As String
   Dim ttxt1 As String
   Dim lPid As Long
   Dim ParentLong As Long
   Dim classInfo As WNDCLASS
   Dim dStyles1 As Long
 
 If mousemove = True Then
    mnuSaveInfo.Enabled = True
    mnuIcon.Enabled = True
    'gets the cursor position
    r = GetCursorPos(cursorPos1)
    'various functions to get information about the window under the cursor
    hWnd1 = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    r = GetClassName(hWnd1, sztext, 100)
    hWnd2 = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    p = GetWindowText(hWnd2, windowname, 100)
    hwnd3 = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    ParentLong = GetParent(hwnd3)
    Call GetWindowRect(hWnd1, WndRECT)
    v& = GetDC(hWnd1)
    
    dStyles1 = GetWindowLong(hWnd1, GWL_STYLE)
    frmStyle.txtDec.Text = dStyles1
    
    'function to get the pixel color under the mouse
    Call GetPixel1(Label6, Label12)
   
    sWnd = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
    Call GetWindowThreadProcessId(sWnd, lPid)
    'function to get the window styles
    Call GetStyles(sWnd, wStyles)
    frmStyle.lblStyle.Text = wStyles
  
    
    'clears the screenshot box so a new image can be put in
    frmDC.Picture2.Picture = Nothing
    'paint screenshot into picturebox
    Call BitBlt(frmDC.Picture2.hDC, 0, 0, 600, 500, v&, 0, 0, vbSrcCopy)
  
    'release screenshot from memory
    
    Call ReleaseDC(hWnd1, v&)

    h2Icon = GetClassLong(hWnd1, GCL_HICON)
    picIcon.Picture = Nothing
    Call DrawIcon(picIcon.hDC, 1, 1, h2Icon)
    
              'function to get unmasked text
              ttxt = Space(100)
              errval = GetCursorPos(cursorPos1)
              thwnd = WindowFromPoint(cursorPos1.X, cursorPos1.Y)
              errval = SendMessage(thwnd, WM_GETTEXT, ByVal TXT_LEN, ByVal ttxt)
              ttxt = RTrim(ttxt)
              
              
              
              ttxt1 = Space(255)
              eerv1 = SendMessage(ParentLong, WM_GETTEXT, ByVal TXT_LEN, ByVal ttxt1)
              ttxt1 = RTrim(ttxt1)
             
    ParentWindow = Space(255)
    Parent = GetParent(hWnd1)
    parent_ = GetClassName(Parent, ParentWindow, 255)


    
    
    'set all the textboxes
    txtUM.Text = ttxt
    txtHandle.Text = hWnd1
    txtParent.Text = ParentLong
    txtClass.Text = sztext
    Text1.Text = txtUM.Text
    txtPText.Text = ttxt1
    frmStyle.txtRect.Text = "Top: " & WndRECT.Top & "    Bottom: " & WndRECT.Bottom & vbCrLf & "Left: " & WndRECT.Left & "    Right: " & WndRECT.Right
    txtPID.Text = Hex(lPid)
    txtPos.Text = "X:  " & cursorPos1.X & "    Y:  " & cursorPos1.Y
    

    txtPClass.Text = ParentWindow

    

    
End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'when the mouse is released after dragging we tell the program to simulate the pointer going back to the box
Picture1.Picture = LoadResPicture(102, vbResIcon)
Form1.MousePointer = 0
'tell the program to stop getting information about the windows
mousemove = False
End Sub


Private Function TextRO(textbx As TextBox)
'function to set the textboxes to readonly
a = SendMessage(textbx.hwnd, EM_SETREADONLY, 1, 0)
End Function






'function to get the pixel color under the mouse
Public Function GetPixel1(PaintLabel As TextBox, DisplayLabel As TextBox)

Dim cursorPos1 As POINTAPI ' variable for the cursor position
Dim hwndWindow As String ' variable for the cursor positions hwnd
Dim wndDC1 As String ' variable for the hwnd's DC
Dim dcPixel As String ' variable for the DC's pixel



a = GetCursorPos(cursorPos1) ' get the current cursor position
hwndWindow = WindowFromPoint(cursorPos1.X, cursorPos1.Y) ' from the cursor position, get the hwnd under it
wndDC1 = GetDC(hwndWindow) ' get DC from the currents hwnd
dcPixel = GetPixel(wndDC1, cursorPos1.X, cursorPos1.Y) ' get the pixel under the mouse from the DC



Dim blue&, green&, red&, colour& ' variabled to change decimal format into RGB




colour& = dcPixel ' set color variable to decimal format

If Len(Str(dcPixel)) = 2 Then ' checks if dcpixel is incorrect format
    Exit Function ' if so, exit the sub without drawing color
Else ' if format is valid, continue

blue& = Int(colour& / 65536) ' function to get the blue
green& = Int((colour& - (65536 * blue&)) / 256) ' function to get the green
red& = colour& - (blue& * 65536) - (green& * 256) ' function to get the red

colour& = RGB(red&, green&, blue&) ' set final RGB format

PaintLabel.BackColor = colour& ' paint dcPixel in RGB format to the color label
DisplayLabel.Text = red & "," & green & "," & blue

End If ' stop asking questions, lol

Call DeleteDC(wndDC)


End Function

Private Sub Text1_Change()
'just little bug fix here. this makes sure that the user is typing in the info, and not ultraspy itself

If mousemove Then
    Exit Sub
Else
Call SendMessage(txtHandle.Text, WM_SETTEXT, ByVal CLng(0), ByVal Text1.Text)
End If
End Sub




Private Sub Timer1_Timer()

i1 = i1 + 5

If i1 = 50 Then
Timer1.Enabled = False
Else
Call FlashWindow(txtHandle.Text, 1)
End If
End Sub
