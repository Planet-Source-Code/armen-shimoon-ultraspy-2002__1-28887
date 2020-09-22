VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shimoon Update"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox frame1 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   2400
      ScaleHeight     =   3135
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.Label Label3 
         Caption         =   "   To begin the update, press next."
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "   This wizard will check a remote server for updates, and if there are any available, will give you the choice to download them."
         Height          =   735
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "  This wizard will help you through the process of updating the Shimoon Technologies software."
         Height          =   615
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   3855
      End
   End
   Begin InetCtlsObjects.Inet net 
      Left            =   3360
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox frame2 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   2400
      ScaleHeight     =   3615
      ScaleWidth      =   4335
      TabIndex        =   5
      Top             =   120
      Width           =   4335
      Begin VB.Label lblStat 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   3120
         Width           =   4095
      End
      Begin VB.Label lblCheck 
         Caption         =   "    The wizard is now connecting to the server and checking for updates..."
         Height          =   2895
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next -->"
      Height          =   375
      Left            =   5520
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "frmUpdate.frx":1CFA
      Top             =   0
      Width           =   2250
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nVer As Single
Dim nMsg As String
Dim nURL As String
Dim b() As Byte

Private Sub cmdCancel_Click()
Unload frmUpdate

End Sub

Private Sub cmdDone_Click()
Unload frmUpdate
End Sub

Private Sub cmdNext_Click()
frame1.Visible = False
frame2.Visible = True

On Error GoTo 10
lblStat.Caption = "Connecting..."
nVer = net.OpenURL("necserver.wox.org/nver.txt")
lblStat.Caption = "Downloading update information..."
nMsg = net.OpenURL("necserver.wox.org/nmsg.txt")
nURL = net.OpenURL("necserver.wox.org/nurl.txt")
lblStat.Caption = "Done downloading information."

Call CheckInfo

Exit Sub

10: lblCheck.Caption = "   Error while processing request."

End Sub

Function CheckInfo()
If UVersion < nVer Then
On Error GoTo 10
lblCheck.Caption = "   Your version is currently outdated. Your version is " & UVersion & ", the new version is " & nVer & "."
lblStat.Caption = "Downloading update file..."
b() = net.OpenURL("http://necserver.wox.org/update.exe", icByteArray)

Open "c:\shimtemp.exe" For Binary Access Write As #1
Put #1, , b()
Close #1

Call Shell("c:\shimtemp.exe", vbNormalFocus)
Call subExit

Else
lblCheck.Caption = "   You have the newest version. Thank you.  (version " & nVer & ")"
End If

lblCheck.Caption = lblCheck.Caption & vbCrLf & vbCrLf & nMsg

cmdNext.Visible = False
cmdDone.Visible = True

Exit Function

10: MsgBox Err.Description, vbExclamation, "Update Error"
End Function






Function subExit()

Dim OnTop As Boolean
Dim fWidth As String
Set fD = Nothing

If Form1.mnuAOT.Checked = True Then
    OnTop = True
Else
    OnTop = False
End If
    
If Form1.mnuSmall.Checked = True Then
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
End Function
