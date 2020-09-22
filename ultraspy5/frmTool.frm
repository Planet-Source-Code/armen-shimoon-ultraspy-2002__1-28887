VERSION 5.00
Begin VB.Form frmTool 
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   840
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Height = Form1.Height
End Sub
