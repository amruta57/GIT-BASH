VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmHomeLoanDetails 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form2"
   ScaleHeight     =   9030
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   10095
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "SEARCH"
         Height          =   615
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1575
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
         Height          =   405
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Customer Name"
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
         Left            =   840
         TabIndex        =   5
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Cancel"
      Height          =   615
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   690
      Left            =   5760
      Top             =   8040
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1217
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bank Application\bank_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bank Application\bank_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from home_loan"
      Caption         =   ""
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Home Loan Details"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   570
      Left            =   4125
      TabIndex        =   6
      Top             =   840
      Width           =   4395
   End
End
Attribute VB_Name = "frmHomeLoanDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MDIForm1.Show
Me.Hide


End Sub

Private Sub Command4_Click()

If Text1.Text = "" Then
        MsgBox "Please enter name ", vbExclamation, Title
        Text1.SetFocus
        
    End If

'Adodc2.Recordset.Filter = "customer_name LIKE '" & Me.Text1.Text & "*'"
'Set DataGrid1.DataSource = Adodc2

Adodc2.RecordSource = "SELECT * FROM home_loan where customer_name = '" & (Text1.Text) & "'"
Set DataGrid1.DataSource = Adodc2
Adodc2.Refresh
Adodc2.Caption = Adodc2.RecordSource

End Sub
