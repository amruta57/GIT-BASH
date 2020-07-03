VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "EXIT"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1320
      Top             =   6360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bank Application\bank_db.mdb;Mode=Share Deny None;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bank Application\bank_db.mdb;Mode=Share Deny None;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from login"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Management System"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1740
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()



Adodc1.RecordSource = "Select * from login where username = '" + Text1.Text + "' and password = '" + Text2.Text + "' "

Adodc1.Refresh

If Adodc1.Recordset.EOF Then

MsgBox "Invalid Login!!", vbCritical, "Failed"

Else

MsgBox "Login Success!!", vbSuccess, "Success"

MDIForm1.Show

Unload Me


End If


End Sub

Private Sub Command2_Click()
End
End Sub


