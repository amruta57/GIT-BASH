VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   ForeColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   7440
      Top             =   5160
   End
   Begin VB.Image Image1 
      Height          =   4140
      Left            =   720
      Picture         =   "Splash.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6495
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   5535
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

frmLogin.Show

Unload Me


End Sub
