VERSION 5.00
Begin VB.Form frmADOTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DSN-Less Connection"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSQLStatement 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Type your SQL Statement Here"
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox txtDBName 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Type your Database Name Here"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtServerName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Type your SQL Server Name Here"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Execute SQL"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   1200
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Database"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1275
   End
End
Attribute VB_Name = "frmADOTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' PLEASE E-MAIL DEATHMOON@HOME.COM IF YOU HAVE
' ANY QUESTIONS / SUGGESTIONS ON THIS CODE.


' MUST REFERENCE
' MICROSOFT ACTIVEX DATA OBJECTS 2.1 LIBRARY

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Qy As New ADODB.Command
Dim sql As String
Dim ConnectionString As String

'Const ConnectionString = "driver={SQL Server};" & _
      "server=PlaceServerNameHere;uid=;pwd=;database=DatabaseNameHere"

Private Sub GetConnectionString()
    ConnectionString = "driver={SQL Server};" & _
        "server=" & Me.txtServerName.Text & ";" & _
        "uid=;pwd=;" & _
        "database=" & Me.txtDBName.Text & ";"
End Sub
    

Private Sub Command1_Click()

    GetConnectionString
    
    With cn
      ' Establish DSN-less connection
      .ConnectionString = ConnectionString
      .ConnectionTimeout = 10
      '.Properties("Prompt") = adPromptNever
      ' This is the default prompting mode in ADO.
      .Open
   End With
   
   MsgBox "Connected"
   Me.Command2.Enabled = True
    
End Sub

Private Sub Command2_Click()
    'sql = "SELECT * FROM employee WHERE login_id='deathmoon'"
    sql = Me.txtSQLStatement.Text
    
    rs.Open sql, cn, adOpenDynamic, adLockReadOnly
    
    MsgBox rs.Fields(0) & " " & rs.Fields(1)
    
    rs.Close
    cn.Close
End Sub

