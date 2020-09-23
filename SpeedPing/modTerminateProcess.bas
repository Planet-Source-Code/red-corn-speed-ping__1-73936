Attribute VB_Name = "modTerminateProcess"
Option Explicit

Private Const TH32CS_SNAPPROCESS = &H2
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const MAX_PATH As Integer = &H104
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_TERMINATE As Long = &H1

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
    ByVal lFlags As Long, _
    ByVal lProcessID As Long _
    ) As Long
    
Private Declare Function Process32First Lib "kernel32" ( _
    ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32 _
    ) As Long
    
Private Declare Function Process32Next Lib "kernel32" ( _
    ByVal hSnapShot As Long, _
    uProcess As PROCESSENTRY32 _
    ) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
     ByVal dwDesiredAccess As Long, _
     ByVal bInheritHandle As Boolean, _
     ByVal dwProcessId As Long _
     ) As Long

Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
     ByVal hProcess As Long, _
     ByRef lphModule As Long, _
     ByVal cb As Long, _
     ByRef lpcbNeeded As Long) As Long

Private Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" ( _
     ByVal hProcess As Long, _
     ByVal hModule As Long, _
     ByVal lpFileName As String, _
     ByVal nSize As Long) As Long

Private Declare Sub CloseHandle Lib "kernel32" ( _
    ByVal hPass As Long _
    )

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Sub TerminateAppProcess(AgentFile As String)
    Dim PE As PROCESSENTRY32
    Dim hSnap As Long
    Dim Result As Boolean
    Dim hProcess As Long
    Dim filename As String * MAX_PATH
    Dim retcode As Long
    PE.dwSize = Len(PE)
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    Result = Process32First(hSnap, PE)
    
    Do While Result
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or SYNCHRONIZE Or PROCESS_TERMINATE, False, PE.th32ProcessID)
        If Not EnumProcessModules(hProcess, 0, 0, 0) = 0 Then
            GetModuleFileNameEx hProcess, 0, filename, MAX_PATH
            'MsgBox filename
            If InStr(1, filename, AgentFile, vbTextCompare) > 0 Then
                retcode = TerminateProcess(hProcess, 0&)
                If retcode = 0 Then
                    MsgBox "Error terminating process!"
                End If
            End If
        End If
        CloseHandle hProcess
        Result = Process32Next(hSnap, PE)
    Loop
    CloseHandle hSnap
End Sub

Public Function CountAppProcess(AgentFile As String) As Integer
    Dim PE As PROCESSENTRY32
    Dim hSnap As Long
    Dim Result As Boolean
    Dim hProcess As Long
    Dim filename As String * MAX_PATH
    Dim retcode As Long
    Dim appcount As Integer
    appcount = 0
    PE.dwSize = Len(PE)
    
    hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    Result = Process32First(hSnap, PE)
    
    Do While Result
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or SYNCHRONIZE Or PROCESS_TERMINATE, False, PE.th32ProcessID)
        If Not EnumProcessModules(hProcess, 0, 0, 0) = 0 Then
            GetModuleFileNameEx hProcess, 0, filename, MAX_PATH
            'MsgBox filename
            If InStr(1, filename, AgentFile, vbTextCompare) > 0 Then
                appcount = appcount + 1
            End If
        End If
        CloseHandle hProcess
        Result = Process32Next(hSnap, PE)
    Loop
    CloseHandle hSnap
    CountAppProcess = appcount
End Function


