VERSION 5.00
Begin VB.UserControl CabinSmall 
   BackColor       =   &H80000007&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   LockControls    =   -1  'True
   ScaleHeight     =   1380
   ScaleWidth      =   2760
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   180
      Top             =   540
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      Height          =   270
      Left            =   2175
      Shape           =   3  'Circle
      Top             =   360
      Width           =   510
   End
   Begin VB.Shape LED 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FF80&
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      Height          =   210
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   390
      Width           =   315
   End
   Begin VB.Label lblAt 
      BackStyle       =   0  'Transparent
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   6
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   210
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "USD 0.0000"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   30
      TabIndex        =   5
      Top             =   930
      Width           =   1380
   End
   Begin VB.Label lblRate 
      BackStyle       =   0  'Transparent
      Caption         =   "USD 0.0200"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   900
      TabIndex        =   4
      Top             =   375
      UseMnemonic     =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblCabinNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   75
      TabIndex        =   3
      Top             =   195
      Width           =   750
   End
   Begin VB.Label lblDestinationName 
      BackStyle       =   0  'Transparent
      Caption         =   "Singapore Mobile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   690
      TabIndex        =   2
      Top             =   135
      UseMnemonic     =   0   'False
      Width           =   1845
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1560
      TabIndex        =   1
      Top             =   945
      Width           =   975
   End
   Begin VB.Label lblNumber 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   0
      Top             =   660
      UseMnemonic     =   0   'False
      Width           =   2565
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H0000C000&
      BorderWidth     =   5
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   2655
   End
End
Attribute VB_Name = "CabinSmall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim boothstate As Boolean

Private startTime As Date
Private duration As Integer

'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Event xEvent() 'MappingInfo=UserControl,UserControl,-1,Click

Enum cbnStateSmall
    stsConnecting
    stsProgressing
    stsConnected
    stsDisconnected
    stsEnd
    stsInvalid
End Enum

Private Enum hkHookTypeSmall
    hksOnHook
    hksOffHook
    hksHookFlash
End Enum
'Default Property Values:
'Const m_def_DestinationName = 0
'Const m_def_DestinationRate = 0
'Property Variables:
'Dim m_DestinationName As Variant
'Dim m_DestinationRate As Variant

Dim sCurrentState As cbnStateSmall

Private Sub OpenDB()
    Set cn = New ADODB.Connection
    cn.Open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " & App.Path & "\callshop.cfg"
       
End Sub

Private Sub CloseDB()
    cn.Close
    
    Set cn = Nothing
    Set rs = Nothing
    
End Sub

Private Sub imgOffHook_Click()
    Label3_Click
End Sub

Private Sub imgOnHook_Click()
    Label3_Click
End Sub

Private Sub Label3_Click()
    If boothstate Then
    
        frmCDR.SetBoothNumber CabinNumber
        DoEvents
        frmCDR.Show vbModal
        
    Else
        MsgBox "Either there is no calls to bill for this booth" & vbCrLf & "Or currently a call is on progress", vbInformation, "Info"
    End If
    
End Sub


Private Sub lblAt_Click()
    Label3_Click
End Sub

Private Sub lblCabinNo_Click()
    Label3_Click
End Sub

Private Sub lblDestinationName_Click()
    Label3_Click
End Sub

Private Sub lblNumber_Click()
    Label3_Click
End Sub

Private Sub lblRate_Click()
    Label3_Click
End Sub

Private Sub lblTime_Click()
    Label3_Click
End Sub

Private Sub tmrTimer_Timer()
Dim diff As String
Dim nowTime As Date
On Error Resume Next

    nowTime = Time
    diff = nowTime - startTime
    duration = duration + 1
    lblTime.Caption = Right("00" & Hour(diff), 2) & ":" & Right("00" & Minute(diff), 2) & ":" & Right("00" & Second(diff), 2)
    Label3.Caption = Left(lblRate, 3) & " " & Round((duration * Mid(lblRate, 5)) / 60, 4)
    
End Sub
'
'Private Sub UserControl_Click()
'    RaiseEvent xEvent
'    RaiseEvent Click
'End Sub
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=shpBorder,shpBorder,-1,BorderWidth
''Public Property Get xProp() As Integer
''    xProp = shpBorder.BorderWidth
''End Property
''
''Public Property Let xProp(ByVal New_xProp As Integer)
''    shpBorder.BorderWidth() = New_xProp
''    PropertyChanged "xProp"
''End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14
'Public Function xMethod() As Variant
'
'End Function

Private Sub UserControl_Click()
    Label3_Click
End Sub

Private Sub UserControl_Initialize()
    OpenDB
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblNumber.Caption = PropBag.ReadProperty("PhoneNumber", "")
    lblDestinationName.Caption = PropBag.ReadProperty("DestinationName", "Singapore - Mobile")
    lblRate.Caption = PropBag.ReadProperty("DestinationRate", "USD 0.0200")
    lblCabinNo.Caption = PropBag.ReadProperty("CabinNumber", "00")
    Label3.Caption = PropBag.ReadProperty("ChargedAmount", "USD 0.0000")
End Sub

Private Sub UserControl_Terminate()
    CloseDB
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("PhoneNumber", lblNumber.Caption, "")
    Call PropBag.WriteProperty("DestinationName", lblDestinationName.Caption, "Singapore - Mobile")
    Call PropBag.WriteProperty("DestinationRate", lblRate.Caption, "USD 0.0200")
    Call PropBag.WriteProperty("CabinNumber", lblCabinNo.Caption, "00")
    Call PropBag.WriteProperty("ChargedAmount", Label3.Caption, "USD 0.0000")
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub ChangeState(ByVal state As cbnStateSmall)
Attribute ChangeState.VB_Description = "Forces a complete repaint of a object."
Dim charge As Single
    
    Select Case state
        Case stsProgressing, stsProgressing:
            'Yellow
            shpBorder.BorderColor = &HC0C0&
            shpBorder.FillColor = &HC0FFFF
            lblCabinNo.ForeColor = &HC0C0&
            boothstate = False
            sCurrentState = stsProgressing
            
            duration = 0
            
           
        Case stsConnected:
            'Blue
            shpBorder.BorderColor = &HC00000
            shpBorder.FillColor = &HFFC0C0
            lblCabinNo.ForeColor = &HC00000
            
            startTime = Time
            tmrTimer.Enabled = True
            duration = 0
            sCurrentState = stsConnected
            
            boothstate = False
        
        Case stsDisconnected
            'Green
            shpBorder.BorderColor = &HC000&
            shpBorder.FillColor = &HC0FFC0
            lblCabinNo.ForeColor = &HC000&
            
            boothstate = False
            duration = 0
            
            sCurrentState = stsDisconnected
            
        Case stsInvalid
            boothstate = False
            shpBorder.BorderColor = &H808080
            shpBorder.FillColor = &HD0D0D0
            lblCabinNo.ForeColor = &H0&
            
            sCurrentState = stsInvalid
                        
        Case stsEnd
            'Red
            shpBorder.BorderColor = &HC0&
            shpBorder.FillColor = &HC0C0FF
            lblCabinNo.ForeColor = &HC0&
                                              
            tmrTimer.Enabled = False
            If Form1.Billtype = "Per Minute" Then
                If duration > 0 Then
                
                    duration = Int((duration + 59) / 60) * 60
                    charge = Round((Mid(DestinationRate, 5) / 60) * duration, 4)
                    Label3 = Left(Label3, 3) & " " & charge
                    'lblTime.Caption = Right("00" & Hour(diff), 2) & ":" & Right("00" & Minute(diff), 2) & ":" & Right("00" & Second(diff), 2)
                    lblTime.Caption = Right("00" & duration \ 3600, 2) & ":" & Right("00" & duration \ 60, 2) & ":00"
                    cn.Execute "INSERT INTO RECORDS(Booth, PhoneNumber, Destination, Rate, Duration, Amount) VALUES(" & CabinNumber & ",'" & PhoneNumber & "','" & DestinationName & "'," & Mid(DestinationRate, 5) & "," & duration & "," & Mid(Label3, 5) & ");"
                End If
            Else
                If duration > 0 Then
                    cn.Execute "INSERT INTO RECORDS(Booth, PhoneNumber, Destination, Rate, Duration, Amount) VALUES(" & CabinNumber & ",'" & PhoneNumber & "','" & DestinationName & "'," & Mid(DestinationRate, 5) & "," & duration & "," & Mid(Label3, 5) & ");"
                End If
                
            End If
            
            sCurrentState = stsEnd
            
            boothstate = True
               
    End Select
    
End Sub

Public Sub ChangeHookState(ByVal Hook As hkHookType)

    Select Case Hook
        Case hksOnHook:
            'imgOffHook.Visible = False
            'imgOnHook.Visible = True
            LED.BackColor = &HFF00&
            LED.BorderColor = &H80FF80
       
        Case hksOffHook, hksHookFlash:
            'imgOnHook.Visible = False
            'imgOffHook.Visible = True
            LED.BackColor = &HFF&
            LED.BorderColor = &H8080FF
            
    End Select
End Sub

Public Property Get CurrentState() As String
    CurrentState = sCurrentState
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblNumber,lblNumber,-1,Caption
Public Property Get PhoneNumber() As String
Attribute PhoneNumber.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    PhoneNumber = lblNumber.Caption
End Property

Public Property Let PhoneNumber(ByVal New_PhoneNumber As String)
    lblNumber.Caption() = New_PhoneNumber
    PropertyChanged "PhoneNumber"
End Property
''
'''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'''MappingInfo=lblCabinNo,lblCabinNo,-1,Caption
''Public Property Get CabinNumber() As String
''    CabinNumber = lblCabinNo.Caption
''End Property
''
''Public Property Let CabinNumber(ByVal New_CabinNumber As String)
''    lblCabinNo.Caption() = New_CabinNumber
''    PropertyChanged "CabinNumber"
''End Property
''
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=lblRate,lblRate,-1,Caption
'Public Property Get CabinNumber() As String
'    CabinNumber = lblRate.Caption
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get DestinationName() As Variant
'    DestinationName = m_DestinationName
'End Property
'
'Public Property Let DestinationName(ByVal New_DestinationName As Variant)
'    m_DestinationName = New_DestinationName
'    PropertyChanged "DestinationName"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=14,0,0,0
'Public Property Get DestinationRate() As Variant
'    DestinationRate = m_DestinationRate
'End Property
'
'Public Property Let DestinationRate(ByVal New_DestinationRate As Variant)
'    m_DestinationRate = New_DestinationRate
'    PropertyChanged "DestinationRate"
'End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_DestinationName = m_def_DestinationName
'    m_DestinationRate = m_def_DestinationRate
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDestinationName,lblDestinationName,-1,Caption
Public Property Get DestinationName() As String
Attribute DestinationName.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    DestinationName = lblDestinationName.Caption
End Property

Public Property Let DestinationName(ByVal New_DestinationName As String)
    lblDestinationName.Caption() = New_DestinationName
    PropertyChanged "DestinationName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblRate,lblRate,-1,Caption
Public Property Get DestinationRate() As String
Attribute DestinationRate.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    DestinationRate = lblRate.Caption
End Property

Public Property Let DestinationRate(ByVal New_DestinationRate As String)
    lblRate.Caption() = New_DestinationRate
    PropertyChanged "DestinationRate"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblCabinNo,lblCabinNo,-1,Caption
Public Property Get CabinNumber() As String
Attribute CabinNumber.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    CabinNumber = lblCabinNo.Caption
End Property

Public Property Let CabinNumber(ByVal New_CabinNumber As String)
    lblCabinNo.Caption() = New_CabinNumber
    PropertyChanged "CabinNumber"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label3,Label3,-1,Caption
Public Property Get ChargedAmount() As String
Attribute ChargedAmount.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    ChargedAmount = Label3.Caption
End Property

Public Property Let ChargedAmount(ByVal New_ChargedAmount As String)
    Label3.Caption() = New_ChargedAmount
    PropertyChanged "ChargedAmount"
End Property

