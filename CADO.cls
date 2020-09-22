VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
' CADO Properties
'------------------------------------------------------------
' Note, not all properties are used with each type
' of connection.  Also, more options can be added
' by adding to the ENUM list and Select Case in
' ConnectRemote
' Such as a password protected Access database,
' I never got around to adding a case for that.
'------------------------------------------------------------
Public DataBaseName As String
Public ServerName As String
Public UserName As String
Public Password As String
'------------------------------------------------------------
' This CN object will be my connection obect that
' I can reference later once  a connection is made
'------------------------------------------------------------
Public cn As ADODB.Connection
'------------------------------------------------------------
' Database type constants
'------------------------------------------------------------
Public Enum DATABASE_TYPES
    SQL_SERVER = 0
    MS_ACCESS97 = 1
    MS_ACCESS2K = 2
    REMOTE_SQLDB = 3
    REMOTE_ACCESS_PATH = 4
    REMOTE_DSN = 5
End Enum
'------------------------------------------------------------
' Called to connect to the specified database,
' based on the values in the public properties
'------------------------------------------------------------
Public Function ConnectRemote(Optional dbTYPE As DATABASE_TYPES = 0) As Boolean
On Error GoTo ConnectRemote_Error
    Dim sConnect As String
    '------------------------------------------------------------
    ' Determine database type
    '------------------------------------------------------------
    Select Case dbTYPE
        Case SQL_SERVER
            sConnect = "DRIVER={SQL Server};SERVER=" & Me.ServerName & ";UID=" & Me.UserName & ";PWD=" & Me.Password & ";DATABASE=" & Me.DataBaseName & ";"
        Case MS_ACCESS97
            sConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & Me.DataBaseName & ";"
        Case MS_ACCESS2K
            sConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.DataBaseName & ";"
        Case REMOTE_SQLDB
            sConnect = "Provider=MS Remote;Remote Server=" & Me.ServerName & ";Remote Provider=MSDASQL;DSN=" & Me.DataBaseName & ";UID=" & Me.UserName & ";PWD=" & Me.Password
        Case REMOTE_ACCESS_PATH
            sConnect = "Provider=MS Remote;Remote Server=" & Me.ServerName & ";Remote Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Me.DataBaseName
        Case REMOTE_DSN
            sConnect = "Provider=MS Remote;Remote Server=" & Me.ServerName & ";DSN=" & Me.DataBaseName & ";UID=" & Me.UserName & ";PWD=" & Me.Password
    End Select
    '------------------------------------------------------------
    ' Attempt connection
    '------------------------------------------------------------
    Set cn = New ADODB.Connection
    cn.Open sConnect
    If cn.State = adStateOpen Then
        ConnectRemote = True
    Else
        ConnectRemote = False
    End If
    ConnectRemote = True
ConnectRemote_Exit:
    Exit Function
ConnectRemote_Error:
    MsgBox Err & ":Error in ConnectRemote.  Error Message: " & Err.Description, vbCritical, "Warning"
    ConnectRemote = False
    Resume ConnectRemote_Exit
End Function

