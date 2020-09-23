Attribute VB_Name = "modFindWindow"
'//Find Window
Option Explicit
Private Const GW_HWNDNEXT = 2
Private Const fwp_startswith = 0
Private Const fwp_contains = 1

Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As Any, _
ByVal lpWindowName As Any) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, _
ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Function FindWindowPartial(ByVal titlepart As String, method As Integer) As Long
    Dim hwndtmp As Long
    Dim ret As Long
    Dim titletmp As String

    hwndtmp = FindWindow(0&, 0&) '²Ä¤G­Ó°Ñ¼ÆIf this parameter is NULL, all window names match
    Do Until hwndtmp = 0
        If hwndtmp <> glMyHwnd Then
            If GetParent(hwndtmp) = 0 Then
                titletmp = Space(256)
                ret = GetWindowText(hwndtmp, titletmp, Len(titletmp)) '¶Ç¦^title¦r¦êªºªø«×,¤£§tterminating null character
                If ret <> 0 Then
                    titletmp = Left(titletmp, ret)
                    If InStr(titletmp, titlepart) = 1 Then
                        FindWindowPartial = hwndtmp
                        Exit Do
                    End If
                End If
            End If
        End If
        hwndtmp = GetWindow(hwndtmp, GW_HWNDNEXT)
    Loop
    
End Function
Public Function GetAgentHwnd() As Long
    GetAgentHwnd = FindWindowPartial("NewPingAgent", fwp_startswith)
End Function



