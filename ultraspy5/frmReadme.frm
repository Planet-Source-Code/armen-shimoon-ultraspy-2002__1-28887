VERSION 5.00
Begin VB.Form frmReadme 
   Caption         =   "Readme"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7680
   Icon            =   "frmReadme.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5445
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmReadme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Text1.Text = "UltraSpy 2002 by Shimoon Technologies." & vbCrLf & _
vbCrLf & "What is it?" & vbCrLf & vbCrLf & _
"UltraSpy is a Windows information spying utility. You can retrieve lots of information " & _
"from Windows, and also alter them in certain ways." & vbCrLf & vbCrLf & _
"Who made it?" & vbCrLf & vbCrLf & _
"Armen Shimoon, founder of Shimoon Technologies developed this application. All parts " & _
"were written by Armen Shimoon, except for parts for the Common Dialog written by Bill Bither. (http://www.atalasoft.com)" & vbCrLf & vbCrLf & _
"Bug Notification or Suggestions" & vbCrLf & vbCrLf & _
"Please contact us if there are any bugs in our software. We will make sure we correct the problem and " & _
"release an update. Also, if you have any suggestions, please feel free to send an email to a_shimoon@hotmail.com " & _
"with 'SUGGESTIONS FOR ULTRASPY' as your topic. If your topic is other then that, it will be automatically deleted. " & _
vbCrLf & vbCrLf & "Thank you very much and don't forget to check for updates!"



End Sub

