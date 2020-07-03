Attribute VB_Name = "Module1"
Public recUsers As New ADODB.Recordset

Public con As New ADODB.Connection



Public Sub ConnectMe()


    con.Open "provider = microsoft.jet.oledb.4.0;data source = " & App.Path & "\bank_db.mdb"
        

End Sub
