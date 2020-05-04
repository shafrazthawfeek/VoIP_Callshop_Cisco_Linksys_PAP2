VERSION 5.00
Begin VB.Form frmCurrency 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Settings"
   ClientHeight    =   4095
   ClientLeft      =   6060
   ClientTop       =   4440
   ClientWidth     =   3675
   Icon            =   "fmrCurrency.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRcpt 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   135
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2250
      Width           =   3420
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1440
      Width           =   1515
   End
   Begin VB.TextBox txtXchange 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MaxLength       =   6
      TabIndex        =   1
      Top             =   570
      Width           =   1230
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "fmrCurrency.frx":0442
      Left            =   2025
      List            =   "fmrCurrency.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1005
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   420
      Left            =   2475
      TabIndex        =   4
      Top             =   3600
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2025
      MaxLength       =   3
      TabIndex        =   0
      Top             =   135
      Width           =   1230
   End
   Begin VB.Label txtCharCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2235
      TabIndex        =   11
      Top             =   1995
      Width           =   1320
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Reciept Text:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   10
      Top             =   1965
      Width           =   1320
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Callshop Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   165
      TabIndex        =   8
      Top             =   1500
      Width           =   1710
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   7
      Top             =   660
      Width           =   1710
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   6
      Top             =   1080
      Width           =   1710
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Currency Symbol"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   210
      Width           =   1785
   End
End
Attribute VB_Name = "frmCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Private Sub Command1_Click()
    If Text1 = "" Then Text1 = "USD"
    If txtXchange = "" Or txtXchange = 0 Then txtXchange = 1
    If txtName = "" Then txtName = "My Callshop"
    If txtRcpt = "" Then txtRcpt = "Receipt"
    
       
    Form1.strCurrency = Left(Text1 & "...", 3)
    Form1.Billtype = Combo1.Text
    Form1.Label1 = txtName
    Form1.sngExchange = txtXchange
    Form1.PrintText = txtRcpt
    
    cn.Execute ("Delete * from Settings")
    cn.Execute ("Insert Into Settings (CurrencySymbol, ExchangeRate, Billtype, CallshopName, PrintText) Values('" & Text1.Text & "'," & txtXchange & ", '" & Combo1.Text & "', '" & txtName.Text & "', '" & txtRcpt & "');")
        
    Unload Me
    
End Sub

Private Sub Form_Load()
    OpenDB
    Set rs = cn.Execute("Select * from Settings;")
    Text1 = rs!CurrencySymbol
    txtXchange = rs!ExchangeRate
    txtName = rs!CallshopName
    txtRcpt = rs!PrintText
    
    If rs!Billtype = "Per Minute" Then
        Combo1.ListIndex = 1
    Else
        Combo1.ListIndex = 0
    
    End If
    
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

Private Sub Form_Unload(Cancel As Integer)
    CloseDB
    
End Sub

Private Sub txtRcpt_Change()
    txtCharCount = "(" & Len(txtRcpt) & "/255)"
    
End Sub
