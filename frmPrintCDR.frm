VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPrintCDR 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13320
   LinkTopic       =   "Form2"
   ScaleHeight     =   6075
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   90
      TabIndex        =   0
      Top             =   555
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   4471
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Printed Receipt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   135
      TabIndex        =   1
      Top             =   75
      Width           =   2265
   End
End
Attribute VB_Name = "frmPrintCDR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub PrintCDR(ByVal header As String, ByRef body() As String, ByVal total As String, ByVal count As Integer)
Dim i As Integer
On Error GoTo ErrHandler

'Me.Show

    MSFlexGrid1.Clear
    MSFlexGrid1.FormatString = header

      
    For i = 1 To count
        'Print rs!Booth & vbtab & rs!PhoneNumber & vbtab & rs!Destination & vbtab & rs!Rate & vbtab & rs!duration & vbtab & rs!Amount
        
        MSFlexGrid1.AddItem body(i)
    Next
        
        'MSFlexGrid1.AddItem "12127773456" & vbTab & "United States - New York" & vbTab & "0.4431" & vbTab & "43:44:00" & vbTab & "1.3443"
        MSFlexGrid1.AddItem ""
        MSFlexGrid1.AddItem ""
        MSFlexGrid1.AddItem "Total Amount" & vbTab & total
               
        Me.Height = MSFlexGrid1.Height + 450
        DoEvents
        
        Me.PrintForm
        DoEvents
        
        MsgBox "Data sent to printer", vbInformation + vbOKOnly, "Success"
        
        Unload Me
Exit Sub

    
ErrHandler:
    Select Case Err.Number
        Case 482: MsgBox "There is no Printers in the System!", vbExclamation + vbOKOnly, "Printer Error"
        Case Else: MsgBox "There is some problem with Printing!", vbExclamation + vbOKOnly, "Printer Error"
    End Select
    
    Unload Me

End Sub

