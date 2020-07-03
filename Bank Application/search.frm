VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   9240
      Top             =   8520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
      RecordSource    =   "select * from vehicle_loan"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   9240
      Top             =   7800
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bank Application\bank_db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bank Application\bank_db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from home_loan"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   1320
      TabIndex        =   6
      Top             =   4080
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "SEARCH VEHICLE LOAN"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "SEARCH HOME LOAN"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "search.frx":0000
      Left            =   2760
      List            =   "search.frx":0028
      TabIndex        =   2
      Text            =   "Select month"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Month"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   1320
      TabIndex        =   1
      Top             =   3120
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Loan Records"
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
      Left            =   3315
      TabIndex        =   0
      Top             =   1080
      Width           =   4965
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

MDIForm1.Show

Me.Hide


End Sub

Private Sub Command2_Click()

       Dim startdate As String
       Dim enddate As String
       
    
       If Combo1.Text = "" Then
        
       MsgBox "Please select month ", vbExclamation, Title
       
        
      End If
    
      
       If Combo1.Text = "January" Then
       
       startdate = Format("1/1/2017", "mm/dd/yyyy")
       enddate = Format("1/31/2017", "mm/dd/yyyy")
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "February" Then
        
       startdate = "2/1/2017"
       enddate = "2/28/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "March" Then
        
       startdate = "3/1/2017"
       enddate = "3/31/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
      
       If Combo1.Text = "April" Then
        
       startdate = "4/1/2017"
       enddate = "4/30/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "May" Then
        
       startdate = "5/1/2017"
       enddate = "5/31/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "June" Then
        
       startdate = "6/1/2017"
       enddate = "6/30/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "July" Then
        
       startdate = "7/1/2017"
       enddate = "7/31/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "August" Then
        
       startdate = "8/1/2017"
       enddate = "8/31/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "September" Then
        
       startdate = "9/1/2017"
       enddate = "9/30/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "October" Then
        
        startdate = "10/1/2017"
        enddate = "10/31/2017"
       
       
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "November" Then
        
       startdate = "11/1/2017"
       enddate = "11/30/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Combo1.Text = "December" Then
        
       startdate = "12/1/2017"
       enddate = "12/31/2017"
        
       Adodc1.RecordSource = "SELECT * FROM home_loan WHERE date_homeloan BETWEEN #" & startdate & " and " & enddate & "# "
      
       Adodc1.Refresh
       
       End If
       
       If Adodc1.Recordset.EOF Then
      
       MsgBox "No Record Found!!", vbExclamation, Title
       
      
       Else
       Set DataGrid1.DataSource = Adodc1
       Adodc1.Refresh
       Adodc1.Caption = Adodc1.RecordSource
      
       End If
       


End Sub

Private Sub Command3_Click()

Dim startdate As String
       Dim enddate As String
       
    
       If Combo1.Text = "" Then
        
       MsgBox "Please select month ", vbExclamation, Title
       
        
      End If
    
      
       If Combo1.Text = "January" Then
       
       startdate = Format("1/1/2017", "mm/dd/yyyy")
       enddate = Format("1/31/2017", "mm/dd/yyyy")
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "February" Then
        
       startdate = "2/1/2017"
       enddate = "2/28/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "March" Then
        
       startdate = "3/1/2017"
       enddate = "3/31/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
      
       If Combo1.Text = "April" Then
        
       startdate = "4/1/2017"
       enddate = "4/30/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "May" Then
        
       startdate = "5/1/2017"
       enddate = "5/31/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "June" Then
        
       startdate = "6/1/2017"
       enddate = "6/30/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "July" Then
        
       startdate = "7/1/2017"
       enddate = "7/31/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "August" Then
        
       startdate = "8/1/2017"
       enddate = "8/31/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "September" Then
        
       startdate = "9/1/2017"
       enddate = "9/30/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "October" Then
        
        startdate = "10/1/2017"
        enddate = "10/31/2017"
       
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "November" Then
        
       startdate = "11/1/2017"
       enddate = "11/30/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & "# and #" & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       
       If Combo1.Text = "December" Then
        
       startdate = "12/1/2017"
       enddate = "12/31/2017"
        
       Adodc2.RecordSource = "SELECT * FROM vehicle_loan WHERE date_vehicleloan BETWEEN #" & startdate & " and " & enddate & "# "
      
       Adodc2.Refresh
       
       End If
       If Adodc2.Recordset.EOF Then
      
       MsgBox "No Record Found!!", vbExclamation, Title
       
      
       Else
       Set DataGrid1.DataSource = Adodc2
       Adodc2.Refresh
       Adodc2.Caption = Adodc2.RecordSource
      
       End If
       
      


End Sub

Private Sub Form_Load()
'Call ConnectMe
Adodc1.Refresh
Adodc2.Refresh



End Sub
