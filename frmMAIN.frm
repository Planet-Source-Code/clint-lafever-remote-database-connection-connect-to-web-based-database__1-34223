VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSQL 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmMAIN.frx":0000
      Top             =   600
      Width           =   7335
   End
   Begin MSComctlLib.ListView lvLIST 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdBUTTON 
      Caption         =   "Connect to web based database and display the data"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' Author:  Clint LaFever [lafeverc@hotmail.com]
' Purpose:  This demo is to show how to use ADO to connect
'                to a database found on a web server.  However,
'                some may not be able to execute this example
'                if you are behind a firewall that does not let
'                you connect on various ports because I do not
'                run my web server on the default port 80 [long
'                story why].  However, the code is still good
'                and you can always try to connect to some other
'                server/database that you have access to.
'------------------------------------------------------------



'------------------------------------------------------------
' Also, if you wish to take the time, you can find
' more stuff by me at my site:  http://vbasic.iscool.net
'------------------------------------------------------------


Option Explicit
Private Sub cmdBUTTON_Click()
'------------------------------------------------------------
' If you know the structure to the NorthWind database
' found in SQL Server, you can try other SQL statements
' if you wish.  Note that I only gave read only
' access to the login account though so just stick
' with SELECT statements.
'------------------------------------------------------------
    If Me.txtSQL.Text <> "" Then
        lvLIST.ListItems.Clear
        lvLIST.ColumnHeaders.Clear
        ConnectandDisplay
    End If
End Sub
Public Sub ConnectandDisplay()
    On Error GoTo ErrorConnectandDisplay
    Dim rs As ADODB.Recordset, lSQL As String, f As Field
    Dim itm As ListItem, x As Long, DB As CADO
    Dim fNAME As String, lNAME As String
    Static ck As Boolean
    Screen.MousePointer = vbHourglass
    '------------------------------------------------------------
    ' Instanciate my CADO object that does my database
    ' connection work for me.
    '------------------------------------------------------------
    Set DB = New CADO
    With DB
        '------------------------------------------------------------
        ' Give DSN known to the server I am connecting to
        '------------------------------------------------------------
        .DataBaseName = "REMOTEDEMO"
        '------------------------------------------------------------
        ' Give address to web server....note the :123,
        ' that is nothing special.  It is just that my
        ' web server does not use the default port 80 so
        ' I need to specify the custom port for my web
        ' server.
        '------------------------------------------------------------
        .ServerName = "http://68.65.61.25:123"
        .UserName = "Admin"
        .Password = ""
        '------------------------------------------------------------
        ' Attempt to connect, if success then go on.
        '------------------------------------------------------------
        If .ConnectRemote(REMOTE_DSN) = True Then
            lSQL = Me.txtSQL.Text
            '------------------------------------------------------------
            ' Open a recordset
            '------------------------------------------------------------
            Set rs = New ADODB.Recordset
            rs.Open lSQL, DB.cn, adOpenKeyset, adLockOptimistic
            '------------------------------------------------------------
            ' If I got data
            '------------------------------------------------------------
            If rs.EOF = False Then
                '------------------------------------------------------------
                ' Loop each field to build column headers
                '------------------------------------------------------------
                For Each f In rs.Fields
                    lvLIST.ColumnHeaders.Add , , f.Name
                Next
                '------------------------------------------------------------
                ' Loop the data and fill the listview
                '------------------------------------------------------------
                While rs.EOF = False
                    For x = 0 To rs.Fields.Count - 1
                        If x = 0 Then
                            Set itm = lvLIST.ListItems.Add(, , "" & rs.Fields(x).Value)
                        Else
                            itm.SubItems(x) = "" & rs.Fields(x).Value
                        End If
                    Next x
                    rs.MoveNext
                Wend
            End If
            '------------------------------------------------------------
            ' This will only be asked the first time you click
            ' the button and I am assuming here you are going
            ' with my default SQL statement querying my table.
            '  This is just here to show you that you can write
            ' to the table as well.
            '------------------------------------------------------------
            If ck = False Then
                If MsgBox("Would you like to add your name to the table?", vbQuestion + vbYesNo) = vbYes Then
                    fNAME = InputBox("Please Enter Your First Name")
                    lNAME = InputBox("Please Enter Your Last Name")
                    rs.AddNew
                    rs.Fields("DEMO_FIRST_NAME") = fNAME
                    rs.Fields("DEMO_LAST_NAME") = lNAME
                    rs.Update
                    MsgBox "Name added.  Re-run the query to see your name in the table."
                End If
            End If
            ck = True
            Set rs = Nothing
            DB.cn.Close
            Set DB = Nothing
        End If
    End With
    Screen.MousePointer = vbDefault
    Exit Sub
ErrorConnectandDisplay:
    Screen.MousePointer = vbDefault
    MsgBox Err & ":Error in ConnectandDisplay.  Error Message: " & Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub
