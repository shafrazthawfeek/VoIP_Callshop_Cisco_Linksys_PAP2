VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate Management"
   ClientHeight    =   8760
   ClientLeft      =   3465
   ClientTop       =   3240
   ClientWidth     =   11475
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   255
      TabIndex        =   8
      Top             =   180
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   435
      Left            =   8820
      TabIndex        =   6
      Top             =   7815
      Width           =   1260
   End
   Begin VB.TextBox loadpath 
      Height          =   390
      Left            =   330
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   7860
      Width           =   8415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10365
      Top             =   6885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Browse 
      Caption         =   "Browse"
      Height          =   420
      Left            =   8805
      TabIndex        =   3
      Top             =   6915
      Width           =   1290
   End
   Begin VB.TextBox txtratepath 
      Height          =   375
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6945
      Width           =   8415
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5700
      Left            =   255
      TabIndex        =   0
      Top             =   795
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   10054
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   7
      Top             =   8010
      Width           =   7860
   End
   Begin VB.Label Label2 
      Caption         =   "Load New Ratesheet"
      Height          =   255
      Left            =   345
      TabIndex        =   5
      Top             =   7560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Export Ratesheet"
      Height          =   315
      Left            =   315
      TabIndex        =   2
      Top             =   6660
      Width           =   4635
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Browse_Click()
On Error Resume Next
 CommonDialog1.ShowSave
 txtratepath = CommonDialog1.FileName
 
        Set rs = cn.Execute("Select * from selling_rates order by Destination ASC")
        If Not (rs.EOF) And Not (rs.BOF) Then rs.MoveFirst
                
        Open txtratepath For Output As #1
        Label3 = "Exporting Data. Please wait.."
        
        While rs.EOF = False
            Print #1, rs(0) & "," & rs(1) & "," & rs(2)
            DoEvents
            rs.MoveNext
        Wend
        
        Close #1
        
        MsgBox "Export Complete", vbInformation + vbOKOnly, "Done"
        
        Label3 = ""
        txtratepath = ""
 
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim s As String
Dim a

    CommonDialog1.ShowOpen
    loadpath = CommonDialog1.FileName
    DoEvents
    
    If loadpath <> "" Then
    
        Open loadpath For Input As #1
        
        Label3 = "Importing Data. Please wait.."
        
        cn.Execute ("Delete from selling_rates;")
        While Not EOF(1)
            Line Input #1, s
            a = Split(s, ",", , vbBinaryCompare)
            cn.Execute ("insert into Selling_Rates(Pref, Destination, Rate) VALUES(" & a(0) & ",'" & a(1) & "'," & a(2) & ");")
            DoEvents
        Wend
      
        MsgBox "Import complete Complete", vbInformation + vbOKOnly, "Done"
    
    End If
    
    Label3 = ""
    
    loadpath = ""
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub OpenDB()
    Set cn = New ADODB.Connection
    cn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\callshop.cfg"
    
   
End Sub

Private Sub CloseDB()
    cn.Close
    
    Set cn = Nothing
    Set rs = Nothing
    
End Sub

Private Sub Form_Load()
    OpenDB
   
    StartForm

End Sub

Private Sub StartForm()
On Error Resume Next

    If MsgBox("Populate ratesheet? This might take a few minutes", vbYesNo + vbInformation, "Load?") = vbYes Then
    
        Set rs = cn.Execute("Select * from selling_rates order by Destination ASC")
        If Not (rs.EOF) And Not (rs.BOF) Then rs.MoveFirst
        
        Label3 = "Populating Data, Please wait.."
        
        'Browse.Enabled = False
        'txtratepath = "Loading Existing Rates, Please Wait.."
    
        MSFlexGrid1.FormatString = "'       Prefix      '|'                 Destination Name               '|'       Rate       '"
        
        While rs.EOF = False
            MSFlexGrid1.AddItem rs(0) & vbTab & rs(1) & vbTab & rs(2)
            DoEvents
            DoEvents
            rs.MoveNext
        Wend
        
    End If
    'txtratepath = ""
    'Browse.Enabled = True
    
    Label3 = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)

    CloseDB
End Sub
