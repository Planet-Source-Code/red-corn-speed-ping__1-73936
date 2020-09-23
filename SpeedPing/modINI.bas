Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetIniInfo() As Long
 Dim iniFile As String
 
 iniFile = App.Path & "\ping.ini" 'Get key value from ini file
 
 Dim lRetContinuedFailAsDown As Long
 Dim bufContinuedFailAsDown As String * 128
 lRetContinuedFailAsDown = GetPrivateProfileString("ping", "ContinuedFailAsDown", "5", bufContinuedFailAsDown, Len(bufContinuedFailAsDown), iniFile)
 glContinuedFailAsDown = CInt(Trim(Left(bufContinuedFailAsDown, lRetContinuedFailAsDown)))
 
 Dim lRetAgentCount As Long
 Dim bufAgentCount As String * 128
 lRetAgentCount = GetPrivateProfileString("ping", "AgentCount", "2", bufAgentCount, Len(bufAgentCount), iniFile)
 glAgentCount = CInt(Trim(Left(bufAgentCount, lRetAgentCount)))
 
 Dim lRetPingCount As Long
 Dim bufPingCount As String * 128
 lRetPingCount = GetPrivateProfileString("ping", "PingCount", "5", bufPingCount, Len(bufPingCount), iniFile)
 glPingCount = CInt(Trim(Left(bufPingCount, lRetPingCount)))
 
 Dim lRetThreshold As Long
 Dim bufThreshold As String * 128
 lRetThreshold = GetPrivateProfileString("ping", "Threshold", "5", bufThreshold, Len(bufThreshold), iniFile)
 glThreshold = CInt(Trim(Left(bufThreshold, lRetThreshold)))
 
 Dim lRetPingTimeOutHost As Long
 Dim bufPingTimeOutHost As String * 128
 lRetPingTimeOutHost = GetPrivateProfileString("ping", "PingTimeOutHost", "1000", bufPingTimeOutHost, Len(bufPingTimeOutHost), iniFile)
 glPingTimeOutHost = CLng(Trim(Left(bufPingTimeOutHost, lRetPingTimeOutHost)))
 
 Dim lRetPingMaxBurst As Long
 Dim bufPingMaxBurst As String * 128
 lRetPingMaxBurst = GetPrivateProfileString("ping", "PingMaxBurst", "220", bufPingMaxBurst, Len(bufPingMaxBurst), iniFile)
 glPingMaxBurst = CLng(Trim(Left(bufPingMaxBurst, lRetPingMaxBurst)))
 
 Dim lRetPingTimeOutBatch As Long
 Dim bufPingTimeOutBatch As String * 128
 lRetPingTimeOutBatch = GetPrivateProfileString("ping", "PingTimeOutBatch", "6000", bufPingTimeOutBatch, Len(bufPingTimeOutBatch), iniFile)
 glPingTimeOutBatch = CLng(Trim(Left(bufPingTimeOutBatch, lRetPingTimeOutBatch)))
 
 Dim lRetWaitForSingleObject As Long
 Dim bufWaitForSingleObject As String * 128
 lRetWaitForSingleObject = GetPrivateProfileString("ping", "WaitForSingleObject", "5", bufWaitForSingleObject, Len(bufWaitForSingleObject), iniFile)
 glWaitForSingleObject = CLng(Trim(Left(bufWaitForSingleObject, lRetWaitForSingleObject)))
 
 Dim lRetDelayStart  As Long
 Dim bufDelayStart  As String * 128
 lRetDelayStart = GetPrivateProfileString("ping", "DelayStart ", "127", bufDelayStart, Len(bufDelayStart), iniFile)
 glDelayStart = CLng(Trim(Left(bufDelayStart, lRetDelayStart))) '±Ä¥Î½è¼Æ
 
 Dim lRetRefreshCycle  As Long
 Dim bufRefreshCycle  As String * 128
 lRetRefreshCycle = GetPrivateProfileString("ping", "RefreshCycle ", "100", bufRefreshCycle, Len(bufRefreshCycle), iniFile)
 glRefreshCycle = CLng(Trim(Left(bufRefreshCycle, lRetRefreshCycle)))
 
 Dim lRetCycleInterval As Long
 Dim bufCycleInterval As String * 128
 lRetCycleInterval = GetPrivateProfileString("ping", "CycleInterval", "7000", bufCycleInterval, Len(bufCycleInterval), iniFile)
 glCycleInterval = CLng(Trim(Left(bufCycleInterval, lRetCycleInterval)))
 
 
 Dim lRetStatisticsCycle As Long
 Dim bufStatisticsCycle As String * 128
 lRetStatisticsCycle = GetPrivateProfileString("ping", "StatisticsCycle", "5", bufStatisticsCycle, Len(bufStatisticsCycle), iniFile)
 glStatisticsCycle = CLng(Trim(Left(bufStatisticsCycle, lRetStatisticsCycle)))

 
 Dim lRetDebugMode As Long
 Dim bufDebugMode As String * 128
 lRetDebugMode = GetPrivateProfileString("ping", "DebugMode", "1", bufDebugMode, Len(bufDebugMode), iniFile)
 glDebugMode = CBool(Trim(Left(bufDebugMode, lRetDebugMode)))
 
 GetIniInfo = 1 'Success

End Function

Public Function SaveIniInfo() As Long
 Dim iniFile As String
 
 iniFile = App.Path & "\ping.ini" 'Get key value from ini file

 Dim lRetContinuedFailAsDown As Long
 Dim bufContinuedFailAsDown As String
 bufContinuedFailAsDown = Trim(glContinuedFailAsDown)
 lRetContinuedFailAsDown = WritePrivateProfileString("ping", "ContinuedFailAsDown", CStr(bufContinuedFailAsDown), iniFile)
 
 Dim lRetAgentCount As Long
 Dim bufAgentCount As String '¦¹³B¤£«Å§i¦r¦êªø«×¡A¥H§K©I¥sWrite...®É¼g¤J¤Óªø¦r¦ê
 bufAgentCount = Trim(glAgentCount)
 lRetAgentCount = WritePrivateProfileString("ping", "AgentCount", CStr(bufAgentCount), iniFile)

 Dim lRetPingCount As Long
 Dim bufPingCount As String
 bufPingCount = Trim(glPingCount)
 lRetPingCount = WritePrivateProfileString("ping", "PingCount", CStr(bufPingCount), iniFile)
 
 Dim lRetThreshold As Long
 Dim bufThreshold As String
 bufThreshold = Trim(glThreshold)
 lRetThreshold = WritePrivateProfileString("ping", "Threshold", CStr(bufThreshold), iniFile)
 
 Dim lRetPingTimeOutHost As Long
 Dim bufPingTimeOutHost As String
 bufPingTimeOutHost = Trim(glPingTimeOutHost)
 lRetPingTimeOutHost = WritePrivateProfileString("ping", "PingTimeOutHost", CStr(bufPingTimeOutHost), iniFile)
 
 Dim lRetPingMaxBurst As Long
 Dim bufPingMaxBurst As String
 bufPingMaxBurst = Trim(glPingMaxBurst)
 lRetPingMaxBurst = WritePrivateProfileString("ping", "PingMaxBurst", CStr(bufPingMaxBurst), iniFile)
 
 Dim lRetPingTimeOutBatch As Long
 Dim bufPingTimeOutBatch As String
 bufPingTimeOutBatch = Trim(glPingTimeOutBatch)
 lRetPingTimeOutBatch = WritePrivateProfileString("ping", "PingTimeOutBatch", CStr(bufPingTimeOutBatch), iniFile)
 
 Dim lRetWaitForSingleObject As Long
 Dim bufWaitForSingleObject As String
 bufWaitForSingleObject = Trim(glWaitForSingleObject)
 lRetWaitForSingleObject = WritePrivateProfileString("ping", "WaitForSingleObject", CStr(bufWaitForSingleObject), iniFile)
 
 Dim lRetDelayStart  As Long
 Dim bufDelayStart  As String
 bufDelayStart = Trim(glDelayStart)
 lRetDelayStart = WritePrivateProfileString("ping", "DelayStart ", CStr(bufDelayStart), iniFile)
 
 Dim lRetRefreshCycle  As Long
 Dim bufRefreshCycle  As String
 bufRefreshCycle = Trim(glRefreshCycle)
 lRetRefreshCycle = WritePrivateProfileString("ping", "RefreshCycle ", CStr(bufRefreshCycle), iniFile)
 
 Dim lRetCycleInterval As Long
 Dim bufCycleInterval As String
 bufCycleInterval = Trim(glCycleInterval)
 lRetCycleInterval = WritePrivateProfileString("ping", "CycleInterval", CStr(bufCycleInterval), iniFile)
 
 Dim lRetStatisticsCycle As Long
 Dim bufStatisticsCycle As String
 bufStatisticsCycle = Trim(glStatisticsCycle)
 lRetStatisticsCycle = WritePrivateProfileString("ping", "StatisticsCycle", CStr(bufStatisticsCycle), iniFile)
 
 Dim lRetDebugMode As Long
 Dim bufDebugMode As String
 bufDebugMode = Trim(glDebugMode)
 lRetDebugMode = WritePrivateProfileString("ping", "DebugMode", CStr(bufDebugMode), iniFile)
 
 SaveIniInfo = 1 'Success

End Function






