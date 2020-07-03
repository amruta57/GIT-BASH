VERSION 5.00
Begin VB.Form frmomeLoan 
   BackColor       =   &H008080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home Loan"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   13020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "BACK"
      Height          =   975
      Left            =   5160
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
End
Attribute VB_Name = "frmomeLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Me.Hide

MDIForm1.Show

End Sub
