VERSION 5.00
Begin VB.Form frmDC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Device Context"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Device Context"
      Height          =   4425
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4695
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Picture"
         Height          =   465
         Left            =   3510
         TabIndex        =   6
         Top             =   810
         Width           =   1005
      End
      Begin VB.CommandButton cmdSaveIcon 
         Caption         =   "Sa&ve Icon"
         Height          =   465
         Left            =   3510
         TabIndex        =   5
         Top             =   270
         Width           =   1005
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
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
         Height          =   645
         Left            =   180
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   360
         Width           =   645
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
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
         Height          =   2715
         Left            =   180
         ScaleHeight     =   2655
         ScaleWidth      =   4275
         TabIndex        =   1
         Top             =   1530
         Width           =   4335
      End
      Begin VB.Label Label13 
         Caption         =   "To extract an icon, simply drag the desired application or library over the picturebox to the left."
         Height          =   825
         Left            =   1080
         TabIndex        =   4
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "Screenshot"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   1260
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IconSelect As Boolean

Private Sub cmdSave_Click()


 
'saves the screenshot
 
If Form1.txtHandle.Text <> "" Then
    
    
    On Error GoTo errorhandler
    With fD
        .flags = OFN_OVERWRITEPROMPT

        .hwnd = Me.hwnd
        .DefaultExt = ".bmp"
        .CancelError = False
        .ShowSave

    End With

    SavePicture picIcon.Image, fD.Filename
    MsgBox "Picture saved to " & fD.Filename, vbInformation, "Saved"
Else
    Exit Sub
End If
Exit Sub

Exit Sub
errorhandler:
    MsgBox Err.Description
End Sub

Private Sub cmdSaveIcon_Click()
'saves our icon we extracted

If (IconSelect = False) Then Exit Sub
On Error GoTo errorhandler
    With fD
        .flags = OFN_OVERWRITEPROMPT

        .hwnd = Me.hwnd
        .DefaultExt = ".ico"
        .CancelError = False
        .ShowSave

    End With



SavePicture picIcon.Image, fD.Filename
MsgBox "Icon saved to " & fD.Filename, vbInformation, "Saved"

Exit Sub
errorhandler:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
Me.Icon = LoadResPicture(101, vbResIcon)
IconSelect = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Me.Hide
Cancel = True
End Sub

Private Sub picIcon_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Filename
Dim FileIcon As Long

'gets the filename of the file dropped onto the picturebox then loads the icon
If (Data.GetFormat(vbCFFiles) = True) Then
    For Each Filename In Data.Files
        FileIcon = ExtractIcon(App.hInstance, Filename, 0)
        frmDC.picIcon.Picture = Nothing
        Call DrawIcon(frmDC.picIcon.hDC, 1, 1, FileIcon)
        IconSelect = True
    Next Filename
End If
End Sub
