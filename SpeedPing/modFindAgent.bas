Attribute VB_Name = "modFindAgent"
Option Explicit

Private Const GW_HWNDNEXT = 2

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwprocessid As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

' Return the window handle for an instance handle.
Public Function InstanceToHwnd(ByVal target_pid As Long) As Long
    Dim test_hwnd As Long
    Dim test_pid As Long
    Dim test_thread_id As Long

    ' Get the first window handle.
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)

    ' Loop until we find the target or we run out
    ' of windows.
    Do While test_hwnd <> 0
        ' See if this window has a parent. If not,
        ' it is a top-level window.
        If GetParent(test_hwnd) = 0 Then
            ' This is a top-level window. See if
            ' it has the target instance handle.
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)

            If test_pid = target_pid Then
                ' This is the target.
                InstanceToHwnd = test_hwnd
                Exit Do
            End If
        End If

        ' Examine the next window.
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Function ProcIDFromWnd(ByVal hwnd As Long) As Long
   Dim idProc As Long
  
   ' Get PID for this HWnd
   GetWindowThreadProcessId hwnd, idProc
   ProcIDFromWnd = idProc
End Function
      
Public Function GetWinHandle(hInstance As Long) As Long
   Dim tempHwnd As Long
      
   ' Grab the first window handle that Windows finds:
   tempHwnd = FindWindow(ByVal 0&, ByVal 0&)
   
   ' Loop until you find a match or there are no more window handles:
   Do Until tempHwnd = 0
      ' Check if no parent for this window
      If GetParent(tempHwnd) = 0 Then
         ' Check for PID match
         If hInstance = ProcIDFromWnd(tempHwnd) Then
            ' Return found handle
            GetWinHandle = tempHwnd
            ' Exit search loop
            Exit Do
         End If
      End If
   
      ' Get the next window handle
      tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT)
   Loop

End Function

