Attribute VB_Name = "Custom"
'Customization Module for - Generic
Option Explicit

Private expiry_date As Date
Private cdrurl As String
Private formcaption As String

Private compact_booth_height As Integer
Private compact_booth_width As Integer
Private standard_booth_height As Integer
Private standard_booth_width As Integer

Enum bType
    bStandard
    bCompact
End Enum

Enum gType
    gHeight
    gWidth
End Enum


Private Sub init_settings()
    'License Expiry Date
    expiry_date = DateSerial(2013, 12, 31)
    
    'CDR Page URL
    cdrurl = "http://account.checkcdr.com"
    
    'MDI Parent Caption
    formcaption = "cdrpanel.com Callshop. Expires in: " & DateDiff("d", Now, expiry_date) & " Days."
    
    compact_booth_height = 1600 ' 1275
    compact_booth_width = 3000 ' 2655
    standard_booth_height = 2350 ' 1845
    standard_booth_width = 5750 ' 5175
    
    
End Sub

Public Function boothGeometry(ByVal b_Type As bType, g_Type As gType) As Integer
'btype: 0-Standard, 1-Compact
'gtype: 0-height, 1:width

Dim out As Integer

Select Case b_Type:
    Case bStandard
        Select Case g_Type
            Case gHeight: out = standard_booth_height
            Case gWidth: out = standard_booth_width
        End Select
        
    Case bCompact
        Select Case g_Type
            Case gHeight: out = compact_booth_height
            Case gWidth: out = compact_booth_width
        End Select
        
End Select

    boothGeometry = out

End Function


Public Function MDI_Caption() As String
    MDI_Caption = formcaption

End Function


Public Function CDR_URL() As String
    CDR_URL = cdrurl
    
End Function


Public Function expiry() As Date
    expiry = expiry_date
    
End Function

Public Function Authorized_Proxy(ByVal msg As String) As Integer
Dim pMessege As String
Dim auth_proxies As Integer

    pMessege = msg
    
    auth_proxies = InStr(9, pMessege, "." & "cdrp" & "anel" & ".com", vbTextCompare)
    auth_proxies = auth_proxies + InStr(9, pMessege, "69" & "." & "41" & "." & "186" & "." & "171", vbTextCompare)
    auth_proxies = auth_proxies + InStr(9, pMessege, "69" & "." & "41" & "." & "186" & "." & "172", vbTextCompare)
    auth_proxies = auth_proxies + InStr(9, pMessege, "69" & "." & "41" & "." & "186" & "." & "173", vbTextCompare)

    Authorized_Proxy = auth_proxies

End Function

Sub main()
    init_settings
    
    Load MDIForm1
    MDIForm1.Show
    
End Sub

