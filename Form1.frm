VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Main Interface"
   ClientHeight    =   6795
   ClientLeft      =   1470
   ClientTop       =   2550
   ClientWidth     =   9480
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDebug 
      Caption         =   "Debug"
      Height          =   495
      Left            =   9960
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox canvas 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   5235
      Left            =   240
      ScaleHeight     =   5235
      ScaleWidth      =   6615
      TabIndex        =   8
      Top             =   1560
      Width           =   6615
      Begin VB.VScrollBar VScroll1 
         Height          =   1335
         LargeChange     =   1000
         Left            =   6360
         Max             =   500
         SmallChange     =   500
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox container 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4320
         Left            =   0
         ScaleHeight     =   4320
         ScaleWidth      =   6135
         TabIndex        =   9
         Top             =   0
         Width           =   6135
         Begin CallShop.Cabin Cabin 
            Height          =   1965
            Index           =   0
            Left            =   255
            Top             =   60
            Visible         =   0   'False
            Width           =   5265
            _ExtentX        =   9287
            _ExtentY        =   3466
         End
         Begin CallShop.CabinSmall CabinSmall 
            Height          =   1380
            Index           =   0
            Left            =   405
            Top             =   2130
            Visible         =   0   'False
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   2434
         End
      End
   End
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Billing Configuration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6180
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   570
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Close Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8010
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   555
      Width           =   1515
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Open CDR Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   4320
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   585
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Booth Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2445
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   585
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Rate Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   570
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2850
      TabIndex        =   1
      Top             =   12450
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6030
      Left            =   6150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   7740
      Visible         =   0   'False
      Width           =   6750
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CALLSHOP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   615
      TabIndex        =   7
      Top             =   45
      Width           =   6165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Public strCurrency As String
Public sngExchange As Single
Public Billtype As String
Public PrintText As String

Public checkdomain As Boolean

Dim sBooths(1 To 16, 0 To 7) As Byte
Dim sIP(1 To 16) As String
Dim rows As Byte, cols As Byte
Dim boothType As Byte '0: Standard, 1: Compact

Public CDRBooth As Integer

Public cabin1 As Object

Private Sub canvas_Resize()
    container.Width = canvas.Width
    'container.Height = canvas.Height
    
    VScroll1.Top = 0
    VScroll1.Left = canvas.Width - VScroll1.Width
    VScroll1.Height = canvas.Height
    
    Select Case boothType
    Case 0: If ((rows * boothGeometry(bStandard, gHeight)) + container.Top + 50) > canvas.Height Then VScroll1.Visible = True Else VScroll1.Visible = False
    Case 1: If ((rows * boothGeometry(bCompact, gHeight)) + container.Top + 50) > canvas.Height Then VScroll1.Visible = True Else VScroll1.Visible = False
    End Select

End Sub

Private Sub Command1_Click()
    Text1 = ""
End Sub

Private Sub Command2_Click()
    Load frmConfig
    frmConfig.Show
End Sub

Private Sub Command3_Click()
    Load frmBooth
    frmBooth.Show
End Sub

Private Sub Command4_Click()
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & CDR_URL()
End Sub


Private Sub Command5_Click()
    Form_QueryUnload 0, 4
End Sub

Private Sub Command6_Click()
    frmCurrency.Show vbModal
End Sub

Private Sub cmdDebug_Click()
Dim msg As String

     msg = "Canvas: " & canvas.Height & vbCrLf & _
        "Container: " & container.Height & vbCrLf & _
        "Rows: " & rows & vbCrLf & _
        "Cols: " & cols & vbCrLf & _
        "Scroll: " & VScroll1.Visible
        
    MsgBox msg
    
    Clipboard.SetText msg


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If Shift = 0 Then
    Select Case KeyCode
        Case 33, 34: If VScroll1.Visible Then VScroll1.SetFocus
    
    End Select
    End If
    
End Sub

Private Sub Form_Load()
    
    Me.Hide
    'VScroll1.Visible = False
    
    strCurrency = "USD"
    Billtype = "Per Second"
    sngExchange = 1
       
    InitSocket
    'initBooths
    OpenDB
    
    Set rs = cn.Execute("Select * from Settings;")
    
    strCurrency = rs!CurrencySymbol
    sngExchange = rs!ExchangeRate
    Label1 = rs!CallshopName
    Billtype = rs!Billtype
    PrintText = rs!PrintText
        
    'boothType = 1
    'DrawInterface 32
    BoothConfig
    
    'frmCDR.Show
    Me.Show
    DoEvents
            
End Sub

Private Sub OpenDB()
    Set cn = New ADODB.Connection
    cn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\callshop.cfg"
    
   
End Sub

Private Sub BoothConfig()
Dim x As Integer
Dim a As Integer
Dim b As Integer

    Set rs = cn.Execute("select * from booths where BoothEnabled = True")
    If Not (rs.EOF) And Not (rs.BOF) Then rs.MoveFirst
    
    a = 1
    
    While rs.EOF = False
    
        If a > 1 Then
            For x = 1 To a
                If sIP(x) = rs!IPAddress Then
                    sBooths(x, rs!Port - 1) = rs!BoothNumber
                    b = a + 1
                Else
                    sIP(a) = rs!IPAddress
                    sBooths(a, rs!Port - 1) = rs!BoothNumber
                End If
            Next
        Else
            sIP(a) = rs!IPAddress
            sBooths(a, rs!Port - 1) = rs!BoothNumber
            a = a + 1
            b = a
        End If
        
        rs.MoveNext
        a = b
        
    Wend
    
    rs.Close
    
    Set rs = cn.Execute("Select BoothIsSmall from Settings;")
    If Not (rs.EOF) And Not (rs.BOF) Then
        rs.MoveFirst
        If (rs!BoothIsSmall) = 0 Then
            boothType = 0
        Else
            boothType = 1
        End If
    Else
        boothType = 0
    End If
    
    
    DrawInterface a - 1
    
End Sub


Private Sub CloseDB()
    cn.Close
    
    Set cn = Nothing
    Set rs = Nothing
    
End Sub

Private Sub GetDestinationDetail(ByVal Pref As String, ByRef Name As String, ByRef Rate As Double)
Dim x As Integer
Dim Prefix As String

    Prefix = Pref

    For x = Len(Pref) - 1 To 1 Step -1
    
        Set rs = cn.Execute("SELECT Destination, Rate from Selling_Rates WHERE Pref ='" & Prefix & "'")
        
        If Not (rs.EOF) And Not (rs.BOF) Then
            Name = rs!Destination
            Rate = rs!Rate
            If sngExchange <> 0 Then
                Rate = Rate * sngExchange
            End If
            
            Exit For
            
            'MsgBox Rate
        Else
            Prefix = Left(Prefix, x)
            
        End If
        
    Next
End Sub

Private Sub InitSocket()
On Error GoTo ErrHandler

Dim LocalPort As Integer

With ws
    LocalPort = 514

    If .LocalPort = Empty Then
        .LocalPort = Trim(LocalPort)
        .Bind .LocalPort
    End If
    
End With

Exit Sub

ErrHandler:
    MsgBox "Error! " & Err.Number, vbCritical

End Sub

Private Sub DrawInterface(ByVal Cabins As Byte)
Dim i As Byte, j As Byte, t As Byte
Dim cnt As Byte

    'cols = Cabins Mod 2
    Select Case boothType
        Case 0: cols = Screen.Width \ boothGeometry(bStandard, gWidth)
        Case 1: cols = Screen.Width \ boothGeometry(bCompact, gWidth)
    End Select
    
    If (Cabins Mod cols) = 0 Then
        rows = (Cabins \ cols)
    Else
        rows = (Cabins \ cols) + 1
    End If
    
    Select Case boothType
    Case 0:
        If ((rows * boothGeometry(bStandard, gHeight)) + container.Top + 50) > canvas.Height Then VScroll1.Visible = True Else VScroll1.Visible = False
        container.Height = rows * boothGeometry(bStandard, gHeight)
    Case 1:
        If ((rows * boothGeometry(bCompact, gHeight)) + container.Top + 50) > canvas.Height Then VScroll1.Visible = True Else VScroll1.Visible = False
        container.Height = rows * boothGeometry(bCompact, gHeight)
    End Select
        
    cnt = 0
    t = cols
    
    Select Case boothType
        Case 0: Set cabin1 = Cabin
        Case 1: Set cabin1 = CabinSmall
    End Select
    
    For i = 1 To rows
    
        If i = rows Then t = (Cabins Mod cols)
        If t = 0 Then t = cols
        
        Select Case boothType
        Case 0:
            For j = 1 To t
                cnt = cnt + 1
                Load cabin1(cnt)
                cabin1(cnt).Visible = True
                cabin1(cnt).Top = (i * boothGeometry(bStandard, gHeight)) - boothGeometry(bStandard, gHeight)
                cabin1(cnt).Left = (j * boothGeometry(bStandard, gWidth)) - boothGeometry(bStandard, gWidth) + 250
                cabin1(cnt).CabinNumber = cnt
                cabin1(cnt).ChargedAmount = strCurrency & " 0.0000"
                cabin1(cnt).DestinationName = ""
                cabin1(cnt).DestinationRate = ""
            Next
        Case 1:
            For j = 1 To t
                cnt = cnt + 1
                Load cabin1(cnt)
                cabin1(cnt).Visible = True
                cabin1(cnt).Top = (i * boothGeometry(bCompact, gHeight)) - boothGeometry(bCompact, gHeight)
                cabin1(cnt).Left = (j * boothGeometry(bCompact, gWidth)) - boothGeometry(bCompact, gWidth) + 250
                cabin1(cnt).CabinNumber = cnt
                cabin1(cnt).ChargedAmount = strCurrency & " 0.0000"
                cabin1(cnt).DestinationName = ""
                cabin1(cnt).DestinationRate = ""
            Next
        End Select
        
    Next
        
End Sub

Private Sub initBooths()
    'sIP(1) = "10.0.0.110"
    'sBooths(1, 0) = 1
    'sBooths(1, 1) = 2
    
    'sIP(2) = "10.0.0.3"
    'sBooths(2, 0) = 3
    'sBooths(2, 1) = 4

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode <> 1 Then
        If MsgBox("Are you sure you want to close the application?" & vbCrLf & "All currently progressing calls will be lost", vbInformation + vbYesNo, "Confirm") = vbYes Then
            CloseDB
            Unload Me
            Cancel = 0
            End
        Else
            Cancel = 1
        
        End If
    End If
    
End Sub



Private Sub Form_Resize()
On Error Resume Next
    canvas.Width = Me.ScaleWidth - canvas.Left
    canvas.Height = Me.ScaleHeight - canvas.Top
    VScroll1.Max = container.Height - canvas.Height

End Sub

Private Sub Label1_Click()
'    frmCDR.SetBoothNumber 1
'    frmCDR.Show
End Sub

Private Sub VScroll1_Change()
    VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    container.Top = -VScroll1.Value
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim inData As String
Dim strIP As String
Dim bIndex As Byte

    ws.GetData inData
    strIP = ws.RemoteHostIP

    Text1.SelText = strIP & " - " & inData & vbCrLf
    getIPindex strIP, bIndex
    
   If bIndex < UBound(sIP) Then
        Parser_PAP2 inData, bIndex
    End If

End Sub

Private Sub getIPindex(ByVal IP As String, ByRef Index As Byte)
Dim x As Byte
    
    For x = 1 To UBound(sIP)
        If sIP(x) = IP Then Exit For
    Next
    
    Index = x
    
End Sub


Private Sub Parser_PAP2(ByVal iMsg As String, ByVal bIPIndex As Byte)
Dim sMessege As String
Dim cabinNo As Byte
Dim PhoneNumber As String, sDestName As String, dRate As Double
Static pMessege As String
Dim domainflag As Integer
    
    sMessege = iMsg
    
    'Line OffHook
    If InStr(1, sMessege, "Off Hook", vbBinaryCompare) > 0 Then
        
        cabinNo = sBooths(bIPIndex, Mid(sMessege, 2, 1))
        cabin1(cabinNo).ChangeHookState hkOffHook
                
    End If
    
    'Line OnHook
    If InStr(1, sMessege, "On Hook", vbBinaryCompare) > 0 Then
        
        cabinNo = sBooths(bIPIndex, Mid(sMessege, 2, 1))
        cabin1(cabinNo).ChangeHookState hkOnHook
        
    End If
    
    'Line Hook Flash
    If InStr(1, sMessege, "Hook Flash", vbBinaryCompare) > 0 Then
        
        cabinNo = sBooths(bIPIndex, Mid(sMessege, 2, 1))
        cabin1(0).ChangeHookState hkHookFlash
        
    End If
    
    If InStr(1, pMessege, "Calling:", vbBinaryCompare) > 0 Then
        
                    
            cabinNo = sBooths(bIPIndex, Mid(sMessege, 2, 1))
            cabin1(cabinNo).ChangeState stProgressing
            PhoneNumber = Mid(Mid(pMessege, 9), 1, InStr(9, pMessege, "@", vbBinaryCompare) - 9)
            
            If Left(PhoneNumber, 2) = "00" Then PhoneNumber = Mid(PhoneNumber, 3)
            If Left(PhoneNumber, 3) = "011" Then PhoneNumber = Mid(PhoneNumber, 4)
            
            Debug.Print pMessege & vbCrLf
            Debug.Print Mid(pMessege, 9) & vbCrLf
            Debug.Print InStr(9, pMessege, "@", vbBinaryCompare) - 9
            
         
'        If checkdomain Then
            
            domainflag = Authorized_Proxy(pMessege)
                        
            Select Case domainflag
            Case Is = 0:
            
                PhoneNumber = "INVALID PROVIDER"
                cabin1(cabinNo).PhoneNumber = PhoneNumber
                cabin1(cabinNo).DestinationName = "In" & "val" & "id" & " P" & "ro" & "vid" & "er"
                cabin1(cabinNo).DestinationRate = "0.0000"
                cabin1(cabinNo).ChargedAmount = strCurrency & " 0.0000"
                
                cabin1(cabinNo).ChangeState stInvalid
            
            Case Else:
            
                sDestName = ""
                dRate = 0
                
                cabin1(cabinNo).PhoneNumber = PhoneNumber
        
                GetDestinationDetail PhoneNumber, sDestName, dRate
                cabin1(cabinNo).DestinationName = sDestName
                cabin1(cabinNo).DestinationRate = strCurrency & " " & dRate
                cabin1(cabinNo).ChargedAmount = strCurrency & " 0.0000"

            End Select
'        Else
'
'                sDestName = ""
'                dRate = 0
'
'                Cabin1(cabinNo).PhoneNumber = PhoneNumber
'
'                GetDestinationDetail PhoneNumber, sDestName, dRate
'                Cabin1(cabinNo).DestinationName = sDestName
'                Cabin1(cabinNo).DestinationRate = strCurrency & " " & dRate
'        End If
            
        
    End If
    
    If InStr(1, sMessege, "CC:Remote Resume", vbBinaryCompare) > 0 Then
    
        cabinNo = sBooths(bIPIndex, Mid(pMessege, 2, 1))
        
        If cabin1(cabinNo).CurrentState <> stInvalid Then
            cabin1(cabinNo).ChangeState stConnected
            
        End If
        
    End If
    
    If InStr(1, sMessege, "CC:CallProgress", vbBinaryCompare) > 0 Then
    
        cabinNo = sBooths(bIPIndex, Mid(pMessege, 2, 1))
        
        If cabin1(cabinNo).CurrentState <> stInvalid Then
            cabin1(cabinNo).ChangeState stProgressing
            
        End If
        
    End If
    
    If InStr(1, sMessege, "CC:Failed w/ Calling", vbBinaryCompare) > 0 Then
        
        cabinNo = sBooths(bIPIndex, Mid(pMessege, 2, 1))
        Set rs = cn.Execute("select count(booth) from records where booth = " & cabinNo)
        If rs(0) = 0 Then
            cabin1(cabinNo).ChangeState stDisconnected
        Else
            cabin1(cabinNo).ChangeState stEnd
        End If
        
    End If
    
    
    If InStr(1, sMessege, "]AUD Rel Call", vbBinaryCompare) > 0 Then
        
        cabinNo = sBooths(bIPIndex, Mid(sMessege, 2, 1))
        
        If cabin1(cabinNo).CurrentState <> stInvalid Then
            cabin1(cabinNo).ChangeState stEnd
            
        End If
        
    End If
    
    pMessege = sMessege
           
End Sub

