VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H008080FF&
   Caption         =   "Loan Management System"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   13980
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu Customer 
      Caption         =   "&Customer"
      Begin VB.Menu NewCustomer 
         Caption         =   "New Customer"
      End
      Begin VB.Menu CustomeDetails 
         Caption         =   "Customer Details"
      End
   End
   Begin VB.Menu Loans 
      Caption         =   "&Loans"
      Begin VB.Menu HomeLoan 
         Caption         =   "&Home Loan"
      End
      Begin VB.Menu VehicleLoan 
         Caption         =   "&Vehicle Loan"
      End
   End
   Begin VB.Menu emicalculator 
      Caption         =   "E&MI Calculator"
   End
   Begin VB.Menu searchloans 
      Caption         =   "Search Loans"
   End
   Begin VB.Menu Reports 
      Caption         =   "&Reports"
      Begin VB.Menu HomeLoanReport 
         Caption         =   "&Home Loan Report"
      End
      Begin VB.Menu VehicleLoanReport 
         Caption         =   "&Vehicle Loan Report"
      End
   End
   Begin VB.Menu Aboutus 
      Caption         =   "&About Us"
   End
   Begin VB.Menu exit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CustomeDetails_Click()
frmCustomerDetails.Show

End Sub

Private Sub emicalculator_Click()

Form2.Show

End Sub


Private Sub exit_Click()
End
End Sub

Private Sub HomeLoan_Click()
   
frmHomeLoan.Show

End Sub

Private Sub HomeLoanReport_Click()

DataReport1.Show

End Sub

Private Sub NewCustomer_Click()

frmNewCustomer.Show

'Me.Hide

End Sub


Private Sub searchloans_Click()

Form1.Show


End Sub

Private Sub VehicleLoan_Click()

vehicle.Show

End Sub

Private Sub VehicleLoanReport_Click()

DataReport2.Show



End Sub
