VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewCustomer 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11760
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
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
      RecordSource    =   "select * from customer"
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
   Begin VB.TextBox Text2 
      DataField       =   "dob"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2760
      TabIndex        =   10
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   1800
      TabIndex        =   3
      Top             =   5280
      Width           =   10215
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
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
         Height          =   735
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "CREATE ACCOUNT"
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         DataField       =   "pincode"
         DataSource      =   "Adodc1"
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
         Left            =   5040
         TabIndex        =   35
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text12 
         DataField       =   "city"
         DataSource      =   "Adodc1"
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
         TabIndex        =   33
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text11 
         DataField       =   "street"
         DataSource      =   "Adodc1"
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
         TabIndex        =   31
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text10 
         DataField       =   "address"
         DataSource      =   "Adodc1"
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
         Left            =   1680
         TabIndex        =   29
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Pincode"
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
         Left            =   4200
         TabIndex        =   34
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Left            =   960
         TabIndex        =   32
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Street"
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
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label16 
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
         Left            =   720
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Width           =   5775
      Begin VB.TextBox Text6 
         DataField       =   "phone"
         DataSource      =   "Adodc1"
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
         Left            =   1680
         TabIndex        =   19
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         DataField       =   "qualification"
         DataSource      =   "Adodc1"
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
         TabIndex        =   17
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         DataField       =   "occupation"
         DataSource      =   "Adodc1"
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
         TabIndex        =   15
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         DataField       =   "gender"
         DataSource      =   "Adodc1"
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
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         DataField       =   "name"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   8
         Top             =   690
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Left            =   480
         TabIndex        =   18
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Occupation"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Left            =   720
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DOB"
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
         Left            =   960
         TabIndex        =   11
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Details"
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   6960
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
      Begin VB.TextBox Text9 
         DataField       =   "branch_code"
         DataSource      =   "Adodc1"
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
         Left            =   2400
         TabIndex        =   27
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         DataField       =   "branch_name"
         DataSource      =   "Adodc1"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         DataField       =   "acc_no"
         DataSource      =   "Adodc1"
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
         Left            =   2400
         TabIndex        =   23
         Top             =   1440
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "acc_type"
         DataSource      =   "Adodc1"
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
         ItemData        =   "frmNewCustomer.frx":0000
         Left            =   2400
         List            =   "frmNewCustomer.frx":000A
         TabIndex        =   21
         Text            =   "Select Account"
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Code"
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
         TabIndex        =   26
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         TabIndex        =   24
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Number"
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
         Left            =   480
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Type"
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
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Information"
         BeginProperty Font 
            Name            =   "Californian FB"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      TabIndex        =   9
      Top             =   3030
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Californian FB"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Customer"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   570
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   3315
   End
End
Attribute VB_Name = "frmNewCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

If Text1.Text = "" Then

GoTo Vali

End If
        
If Text2.Text = "" Then
      
GoTo Vali

End If
      
If Text3.Text = "" Then

GoTo Vali

End If
      
If Text4.Text = "" Then
      
GoTo Vali

End If
      
If Text5.Text = "" Then

GoTo Vali

End If
      
If Text6.Text = "" Then
      
GoTo Vali

End If
      
If Combo1.Text = "" Then

GoTo Vali

End If

If Text7.Text = "" Then
  
GoTo Vali

End If
  
If Text8.Text = "" Then

GoTo Vali

End If

If Text9.Text = "" Then

GoTo Vali

End If

If Text10.Text = "" Then

GoTo Vali

End If

If Text11.Text = "" Then

GoTo Vali

End If

If Text12.Text = "" Then

GoTo Vali

End If

If Text13.Text = "" Then
        
GoTo Vali

End If
        

con.Execute "INSERT INTO customer(name, dob, gender, occupation, qualification, phone, address, street, city, pincode, acc_type, acc_no, branch_name, branch_code) VALUES ('" & (Text1.Text) & "','" & (Text2.Text) & "','" & (Text3.Text) & "','" & (Text4.Text) & "','" & (Text5.Text) & "','" & (Text6.Text) & "','" & (Text10.Text) & "','" & (Text11.Text) & "','" & (Text12.Text) & "','" & (Text13.Text) & "','" & (Combo1.Text) & "','" & (Text7.Text) & "','" & (Text8.Text) & "','" & (Text9.Text) & "')"
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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Combo1.Text = ""

Vali: MsgBox "Please enter all data ", vbExclamation, Title

End Sub


Sub Vali()
MsgBox "Please enter all data ", vbExclamation, Title
End Sub


Private Sub Command2_Click()

MDIForm1.Show

Me.Hide

End Sub





Private Sub Form_Load()

Call ConnectMe

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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
Combo1.Text = ""



End Sub
