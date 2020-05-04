VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "cdrpanel.com Callshop"
   ClientHeight    =   8010
   ClientLeft      =   3045
   ClientTop       =   1965
   ClientWidth     =   10305
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Dim saved As Date

    If GetSetting("IE7.0", "GUID", "LR") = "" Then SaveSetting "IE7.0", "GUID", "LR", Date
        
    If DateDiff("d", GetSetting("IE7.0", "GUID", "LR"), Date) < 0 Then
        MsgBox "Please Fix your PC Date/Clock before starting Callshop", vbInformation & vbOKOnly, "Clock Error"
        End
    End If
       
    SaveSetting "IE7.0", "GUID", "LR", Date
    
    saved = GetSetting("IE7.0", "GUID", "LR")
    
    If DateDiff("d", saved, expiry()) <= 0 Then
        MsgBox "This Software license has expired. Please contact your service provider to obtain a new License", vbInformation + vbOKOnly, "License Expired"
        End
    End If
    
    'If DateDiff("d", saved, DateSerial(2007, 9, 10)) <= 0 Then
        'Form1.checkdomain = True
    'Else
        'Form1.checkdomain = False
    'End If
    
    Me.Caption = MDI_Caption
        
    Form1.Show
End Sub

Private Sub MDIForm_Resize()
    Form1.Refresh
    
End Sub
