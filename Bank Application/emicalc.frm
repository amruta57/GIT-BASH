VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMI Calculator"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   7080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "CALCULATE"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMI Calculator"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   570
      Left            =   3435
      TabIndex        =   10
      Top             =   960
      Width           =   3435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Amount"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Rate"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Term (yrs)"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly EMI"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim N As Integer
Dim amt, payment, rate As Double
amt = Val(Text1.Text)
rate = (Val(Text2.Text) / 100) / 12
N = Val(Text3.Text) * 12
payment = Pmt(rate, N, -amt, 0, 0)
Text4.Text = Format(payment, "#,##0.00")

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

End Sub

Private Sub Command3_Click()

Me.Hide

End Sub
