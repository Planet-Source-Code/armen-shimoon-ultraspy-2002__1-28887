VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   2985
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   4485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   2985
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3600
      Top             =   -90
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()


AlwaysOnTop frmSplash, False





End Sub

Private Sub Timer1_Timer()
frmSplash.Hide

Form1.cmdClose.Left = Form1.cmdClose.Left + 50
Form1.Picture1.Picture = LoadResPicture(102, vbResIcon) ' loads the pointer
Form1.Icon = LoadResPicture(101, vbResIcon) ' loads the window icon
Form1.Caption = Chr(85) & Chr(108) & Chr(116) & Chr(114) & Chr(97) & Chr(83) & Chr(112) & Chr(121) & " 2002"
'I developed the Chr idea because this way it will make it hard for hex editors to alter your programs title!

Timer1.Enabled = False
Form1.Show
End Sub
