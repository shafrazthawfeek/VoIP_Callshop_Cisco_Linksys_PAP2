VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCDR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Call Details"
   ClientHeight    =   6600
   ClientLeft      =   2775
   ClientTop       =   2550
   ClientWidth     =   9450
   Icon            =   "frmCDR.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7515
      TabIndex        =   4
      Top             =   5940
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   135
      TabIndex        =   3
      Top             =   5955
      Width           =   1650
   End
   Begin VB.CommandButton cmdPaid 
      Caption         =   "PAID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1980
      TabIndex        =   2
      Top             =   5955
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4725
      Left            =   75
      TabIndex        =   0
      Top             =   165
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2325
      TabIndex        =   1
      Top             =   4965
      Width           =   6975
   End
End
Attribute VB_Name = "frmCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim PrintHead As String
Dim PrintBody(1 To 50, 1 To 5) As String
Dim PrintCount As Integer
Dim PrintTotal As String
Dim AlreadLoaded As Boolean

Private Sub cmdPaid_Click()
    If MsgBox("All records would be cleard", vbOKCancel + vbInformation, "Confirm") = vbOK Then
        cn.Execute "DELETE FROM Records where Booth = " & Form1.CDRBooth
        
        If Form1.cabin1(Form1.CDRBooth).CurrentState = stEnd Then
            Form1.cabin1(Form1.CDRBooth).ChangeState stDisconnected
            Form1.cabin1(Form1.CDRBooth).DestinationName = ""
            Form1.cabin1(Form1.CDRBooth).DestinationRate = ""
            Form1.cabin1(Form1.CDRBooth).ChargedAmount = Left(Form1.cabin1(Form1.CDRBooth).ChargedAmount, 3) & " 0.0000"
            Form1.cabin1(Form1.CDRBooth).PhoneNumber = ""
        End If
        
        AlreadLoaded = False
        Unload Me
    End If
    
End Sub

Private Sub cmdPrint_Click()
    'Load frmPrintCDR
    'frmPrintCDR.Hide
    'frmPrintCDR.PrintCDR PrintHead, PrintBody, PrintTotal, PrintCount
    
Dim i As Integer
Dim yPos As Long
    
Dim PRN As Object
Dim fnt As New StdFont
Dim header() As String

Set PRN = Printer
    
PRN.ScaleMode = 6

header = Split(Form1.PrintText, vbCrLf)
    
fnt.Name = "Arial": fnt.Size = 12: fnt.Bold = True
Set PRN.Font = fnt

PRN.CurrentY = 15

For i = 0 To UBound(header)
    PRN.CurrentX = (PRN.ScaleWidth - PRN.TextWidth(header(i))) / 2
    'PRN.CurrentY = yPos
    PRN.Print header(i)
Next

PRN.Line (10, PRN.CurrentY)-(PRN.ScaleWidth - 10, PRN.CurrentY)
PRN.Print

fnt.Name = "Arial": fnt.Size = 10: fnt.Bold = True
Set PRN.Font = fnt
'1-Number
'2-Destination
'3-Duration
'4-Rate
'5-Amount


    'Print Number & Rate Header
    PRN.CurrentX = 20
    yPos = PRN.CurrentY
    PRN.Print "Description"
       
    'Print Duration Header
    PRN.CurrentY = yPos
    PRN.CurrentX = (PRN.ScaleWidth - PRN.TextWidth("Duration")) / 2
    PRN.Print "Duration"
        
    'Print Amount Header
    PRN.CurrentY = yPos
    PRN.CurrentX = PRN.ScaleWidth - PRN.TextWidth("Amount") - 15
    PRN.Print "Amount"
    
    PRN.Print

fnt.Name = "Arial": fnt.Size = 10: fnt.Bold = False
Set PRN.Font = fnt

i = 0
For i = 1 To PrintCount
    'Print Number & Rate
    PRN.CurrentX = 20
    yPos = PRN.CurrentY
    PRN.Print "[" & i & "] "; PrintBody(i, 1) & " @ " & PrintBody(i, 4) & " / min"
       
    'Print Duration
    PRN.CurrentY = yPos
    PRN.CurrentX = (PRN.ScaleWidth - PRN.TextWidth(PrintBody(i, 3) & " mins")) / 2
    PRN.Print PrintBody(i, 3) & " mins"
        
    'Print Amount
    PRN.CurrentY = yPos
    PRN.CurrentX = PRN.ScaleWidth - PRN.TextWidth(PrintBody(i, 5)) - 15
    PRN.Print PrintBody(i, 5)
    
    'Print Amount
    PRN.CurrentX = 20
    PRN.Print PrintBody(i, 2)
    
    PRN.Print
Next

PRN.Print

fnt.Name = "Arial": fnt.Size = 12: fnt.Bold = True
Set PRN.Font = fnt

'Print Total Label
yPos = PRN.CurrentY
PRN.CurrentX = 20
PRN.Print "Total Amount: "

    'Print Total - Right Align
    PRN.CurrentY = yPos
    PRN.CurrentX = PRN.ScaleWidth - PRN.TextWidth(PrintTotal) - 15
    PRN.Print PrintTotal

PRN.Print

PRN.CurrentX = (PRN.ScaleWidth - PRN.TextWidth("---- End of Receipt ----")) / 2
PRN.Print "---- End of Receipt ----"

Printer.EndDoc

    
End Sub

Private Sub Command1_Click()
    AlreadLoaded = False
    Unload Me
End Sub

Private Sub Command2_Click()

    
End Sub

Private Sub Form_Activate()
If Not AlreadLoaded Then

AlreadLoaded = True
Dim dur As String
Dim total As Currency

    total = 0
    PrintHead = ""
    PrintCount = 0

    Set rs = cn.Execute("SELECT * FROM Records WHERE Booth =" & Form1.CDRBooth & " Order by Id;")
    If Not (rs.EOF) And Not (rs.BOF) Then rs.MoveFirst
    
    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = "   Phone Number  |   Destination   |   Rate   |  Duration  |    Amount   "
    
    PrintHead = "    Phone Number    |         Destination        |   Rate   |  Duration  |    Amount   "
      
    While Not rs.EOF
        'Print rs!Booth & vbtab & rs!PhoneNumber & vbtab & rs!Destination & vbtab & rs!Rate & vbtab & rs!duration & vbtab & rs!Amount
        dur = Right("00" & (rs!duration \ 60), 2) & ":" & Right("00" & (rs!duration Mod 60), 2)
        MSFlexGrid1.AddItem rs!PhoneNumber & vbTab & rs!Destination & vbTab & rs!Rate & vbTab & dur & vbTab & rs!Amount
        total = total + rs!Amount
        
        PrintCount = PrintCount + 1
        'PrintBody(PrintCount) = "(" & PrintCount & ")" & rs!PhoneNumber & " (" & rs!Destination & ") Charged at: " & rs!Rate & ", Spoken for: " & dur & " Mins. Cost: " & rs!Amount
        PrintBody(PrintCount, 1) = rs!PhoneNumber
        PrintBody(PrintCount, 2) = rs!Destination
        PrintBody(PrintCount, 3) = dur
        PrintBody(PrintCount, 4) = rs!Rate
        PrintBody(PrintCount, 5) = rs!Amount
        
        
        
        rs.MoveNext
    Wend
        
    lblTotal.Caption = "Total Payable: " & Form1.strCurrency & " " & Round(total, 4)
    PrintTotal = Form1.strCurrency & " " & Round(total, 4)
End If
End Sub


Private Sub Form_Load()
    OpenDB

End Sub

Public Function SetBoothNumber(ByRef BoothNumber As Integer)
    Form1.CDRBooth = BoothNumber
End Function

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
    AlreadLoaded = False
    CloseDB
End Sub
