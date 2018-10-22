Attribute VB_Name = "ModRBI"
Public db As New ADODB.Connection
Public rs As New ADODB.Recordset
Public un, pw, bn, bc, bs As String
Sub Main()
On Error GoTo dbError
    'unquote the code below to reset
    'SaveSetting App.EXEName, "Config", "Configured", "False"
    Set db = New ADODB.Connection
    db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\dbDatabase.mdb"
    db.Open
    If GetSetting(App.EXEName, "Config", "Configured") = "True" Then
        Load frmSplash
        frmSplash.Show
    Else
        Load frmConfig1
        frmConfig1.Show
    End If
Exit Sub
dbError:
    MsgBox "Unable to connect established to the database, Please try again!", vbExclamation, "Registry of Barangay Inhabitants"
    End
End Sub
