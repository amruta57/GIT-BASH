VERSION 5.00
Begin VB.Form EMICalculator 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "EMI Calculator"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   9480
      TabIndex        =   11
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      Height          =   615
      Left            =   7680
      TabIndex        =   10
      Top             =   5400
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CALCULATE"
      Height          =   615
      Left            =   4800
      TabIndex        =   9
      Top             =   5400
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
      Left            =   4800
      TabIndex        =   8
      Top             =   6720
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
      Left            =   4800
      TabIndex        =   6
      Top             =   3960
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
      Left            =   4800
      TabIndex        =   4
      Top             =   3120
      Width           =   2415
   End
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
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   2415
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
      Left            =   3240
      TabIndex        =   7
      Top             =   6840
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
      Left            =   3120
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
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
      Left            =   3360
      TabIndex        =   3
      Top             =   3240
      Width           =   1335
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
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
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
      Left            =   4395
      TabIndex        =   0
      Top             =   600
      Width           =   3435
   End
End
Attribute VB_Name = "EMICalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

        'Dim principal, years As Integer
        'Dim rate, interest, amount As Single

        'principal = Text1.Text
        'years = Text3.Text
        'rate = Text2.Text

        'If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
        
       ' MsgBox "Enter All Fields", vbSuccess, "Success"
        
       ' Else
        
        
        '    amount = principal * Math.Pow((1 + rate / 100), years)
         '   interest = amount - principal
                
          '  Text4.Text = interest
                
                
    'End If


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
