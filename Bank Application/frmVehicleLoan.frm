VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form VehicleLoan 
   BackColor       =   &H00FFFF80&
   Caption         =   "Vehicle Loan"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
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
      Left            =   3240
      TabIndex        =   27
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox Text13 
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
      Left            =   3240
      TabIndex        =   24
      Top             =   6120
      Width           =   1815
   End
   Begin VB.TextBox Text12 
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
      Left            =   3240
      TabIndex        =   23
      Top             =   5520
      Width           =   1815
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
      Height          =   375
      Left            =   3240
      TabIndex        =   22
      Top             =   4920
      Width           =   1815
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
      Left            =   7200
      TabIndex        =   3
      Top             =   2400
      Width           =   1815
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
      Left            =   7200
      TabIndex        =   2
      Top             =   3120
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "EMI"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   3480
      Top             =   7800
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1005
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=N:\vb projects\Bank Application\bank_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=N:\vb projects\Bank Application\bank_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from home_loan"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Loan Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   10935
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmVehicleLoan.frx":0000
         Left            =   3000
         List            =   "frmVehicleLoan.frx":000A
         TabIndex        =   21
         Text            =   "Select Type"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   2535
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
         Left            =   3000
         TabIndex        =   8
         Top             =   960
         Width           =   1455
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
         TabIndex        =   7
         Top             =   1680
         Width           =   2895
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
         Left            =   3000
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Interest Rate"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Term"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   20
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "(in months)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5040
         TabIndex        =   14
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Model Price"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Loan Amount"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   10
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Loan Amount"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   25
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Loan"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   3075
      TabIndex        =   19
      Top             =   840
      Width           =   4635
   End
End
Attribute VB_Name = "VehicleLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()
EMICalculator.Show

End Sub

Private Sub Text10_Change()

End Sub
