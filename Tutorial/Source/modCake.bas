Attribute VB_Name = "modCake"
Option Explicit
Public cn As ADODB.Connection
Public crx As New CRAXDRT.Application

Public Function OpenDatabase() As Boolean
On Error GoTo checkErr
Set cn = New ADODB.Connection
cn.Provider = "Microsoft.Jet.OLEDB.4.0"
cn.ConnectionString = "Data Source=" & App.Path & "\Data\dbCake.mdb"
cn.Properties("Jet OLEDB:Database Password") = "yuMMy20"
cn.Open
OpenDatabase = True
Exit Function
checkErr:
    MsgBox Err.Description, vbExclamation, Err.Number
End Function
