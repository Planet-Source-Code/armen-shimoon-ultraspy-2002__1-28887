VERSION 5.00
Begin VB.Form frmStyle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Styles"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
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
   ScaleHeight     =   3540
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Style Info"
      Height          =   3345
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4785
      Begin VB.TextBox txtDec 
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   2880
         Width           =   3435
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   465
         Left            =   3870
         TabIndex        =   6
         Top             =   2790
         Width           =   825
      End
      Begin VB.TextBox txtRect 
         Height          =   825
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1710
         Width           =   4455
      End
      Begin VB.TextBox lblStyle 
         Height          =   825
         Left            =   180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   540
         Width           =   4515
      End
      Begin VB.Label Label4 
         Caption         =   "Decimal WS_STYLES"
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   2610
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Rectangle"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "WS_STYLES"
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   1635
      End
      Begin VB.Line Line19 
         BorderColor     =   &H00FFFFFF&
         X1              =   3780
         X2              =   3780
         Y1              =   3330
         Y2              =   2700
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00FFFFFF&
         X1              =   4770
         X2              =   3780
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Line Line17 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   3780
         X2              =   3780
         Y1              =   3330
         Y2              =   2700
      End
      Begin VB.Line Line15 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   4770
         X2              =   3780
         Y1              =   2700
         Y2              =   2700
      End
      Begin VB.Label Label5 
         Height          =   645
         Left            =   3780
         TabIndex        =   1
         Top             =   2700
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Form_Load()
frmStyle.Icon = LoadResPicture(101, vbResIcon)

End Sub

Private Sub lblExStyle_Change()

End Sub

