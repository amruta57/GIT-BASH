VERSION 5.00
Begin VB.Form frmHomeLoan 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home Loan"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   13005
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6480
      Width           =   1455
   End
   Begin VB.TextBox Text11 
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
      Left            =   4200
      TabIndex        =   28
      Top             =   6000
      Width           =   1815
   End
   Begin VB.TextBox Text10 
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
      Left            =   8640
      TabIndex        =   25
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "EMI"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text5 
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
      Left            =   8640
      TabIndex        =   8
      Top             =   2880
      Width           =   2895
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
      Left            =   8640
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1335
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
      Left            =   4200
      TabIndex        =   6
      Top             =   5400
      Width           =   1815
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
      Left            =   4200
      TabIndex        =   5
      Top             =   3600
      Width           =   1815
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   4560
      ScaleHeight     =   510
      ScaleWidth      =   4875
      TabIndex        =   30
      Top             =   7200
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Loan Details"
      Height          =   5175
      Left            =   1680
      TabIndex        =   13
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Text9 
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
         Left            =   2520
         TabIndex        =   4
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6960
         TabIndex        =   9
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   1455
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
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2520
         TabIndex        =   3
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "EMI"
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
         Left            =   1680
         TabIndex        =   27
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   6240
         TabIndex        =   26
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
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
         Left            =   6000
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
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
         Left            =   5880
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label7 
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
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Property Price"
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
         Left            =   840
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "(in months)"
         Height          =   255
         Left            =   4440
         TabIndex        =   18
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Term"
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
         Left            =   1200
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
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
         Left            =   1320
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Left            =   600
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Loan"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   480
      Width           =   2625
   End
End
Attribute VB_Name = "frmHomeLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

con.Execute "INSERT INTO home_loan(customer_name, reference, contact, email, address, loan_amount, loan_term, property_price, interest_rate, date_homeloan, emi) VALUES ('" & (Text3.Text) & "','" & (Text7.Text) & "','" & (Text4.Text) & "','" & (Text5.Text) & "','" & (Text8.Text) & "','" & (Text6.Text) & "','" & (Text9.Text) & "','" & (Text1.Text) & "','" & (Text2.Text) & "', '" & (Text10.Text) & "', '" & (Text11.Text) & "' )"
MsgBox "Record Saved", vbSuccess, "Success"

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""


End Sub

Private Sub Command2_Click()

MDIForm1.Show
Me.Hide

End Sub




Private Sub Command3_Click()
EMICalculator.Show

End Sub

Private Sub Command4_Click()

Dim N As Integer
Dim amt, payment, rate As Double
amt = Val(Text6.Text)
rate = (Val(Text2.Text) / 100) / 12
N = Val(Text9.Text) * 12
payment = Pmt(rate, N, -amt, 0, 0)
Text11.Text = Format(payment, "#,##0.00")

End Sub

Private Sub Command5_Click()

frmHomeLoanDetails.Show
Me.Hide


End Sub

Private Sub Form_Load()


Text10.Text = Date
'Call ConnectMe
Adodc1.Refresh

End Sub

