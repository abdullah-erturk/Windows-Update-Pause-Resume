
' INFO:

' Windows Update Pause / Resume
' Copyright (C) 2024 Abdullah ERTÜRK
' https://github.com/abdullah-erturk/Windows-Update-Pause-Resume

'=====================================
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
  WScript.Quit
End If
'=====================================
Set objShell = CreateObject("WScript.Shell")
Const HKEY_LOCAL_MACHINE = &H80000002
strKeyPath = "SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
strValueName = "PauseFeatureUpdatesStartTime"

Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

On Error Resume Next

intReturnValue = objRegistry.GetStringValue(HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue)
If intReturnValue = 0 Then
	Set objShell = CreateObject("WScript.Shell")
	strKeyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings"
	strValueName1 = "PauseFeatureUpdatesStartTime"
	strValueName2 = "PauseFeatureUpdatesEndTime"
	strValueName3 = "PauseQualityUpdatesStartTime"
	strValueName4 = "PauseQualityUpdatesEndTime"
	strValueName5 = "PauseUpdatesStartTime"
	strValueName6 = "PauseUpdatesExpiryTime"
	strValueName7 = "FlightSettingsMaxPauseDays"

	objShell.RegDelete strKeyPath & "\" & strValueName1
	objShell.RegDelete strKeyPath & "\" & strValueName2
	objShell.RegDelete strKeyPath & "\" & strValueName3
	objShell.RegDelete strKeyPath & "\" & strValueName4
	objShell.RegDelete strKeyPath & "\" & strValueName5
	objShell.RegDelete strKeyPath & "\" & strValueName6
	objShell.RegDelete strKeyPath & "\" & strValueName7

    objShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings\FlightSettingsMaxPauseDays", 35, "REG_DWORD"
    MsgBox "Windows Update Pause Duration changed to default setting (5 weeks)." & vbCR & "" & vbCR & "Windows Update service is enabled.", vbInformation + vbSystemModal, "Win Update Pause / Resume | by Abdullah ERTÜRK"

	Set objShell = CreateObject("WScript.Shell")
	strPowerShellCmd1 = "Stop-Process -Name SystemSettings -Force"
	strPowerShellCmd2 = "Start-Process ms-settings:windowsupdate"
	objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd1 & """", 0, True
	objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd2 & """", 0, True

	Set objShell = Nothing

Else

    PauseWeeks = InputBox("Enter the number of weeks you want to pause Windows Updates:" & vbCR & "" & vbCR & "Example: 10, 20, 30, 40, etc.", "Duration to Pause Updates")

    If IsNumeric(PauseWeeks) Then
        PauseWeeks = CInt(PauseWeeks)
        If PauseWeeks > 0 Then
            FlightSettingsMaxPauseDays = PauseWeeks * 7 

            objShell.RegWrite "HKEY_LOCAL_MACHINE\" & strKeyPath & "\FlightSettingsMaxPauseDays", CInt(FlightSettingsMaxPauseDays), "REG_DWORD"

            objShell.Run "cmd /c net start wuauserv", 0, True

            MsgBox "Added " & PauseWeeks & " weeks (" & FlightSettingsMaxPauseDays & " days) to Windows Updates Pause option.", vbInformation + vbSystemModal, "Win Update Pause / Resume | by Abdullah ERTÜRK"
	
			Set objShell = CreateObject("WScript.Shell")
			strPowerShellCmd1 = "Stop-Process -Name SystemSettings -Force"
			strPowerShellCmd2 = "Start-Process ms-settings:windowsupdate"
			objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd1 & """", 0, True
			objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd2 & """", 0, True
			
			Set objShell = Nothing

        End If
    Else
        MsgBox "Please enter a valid number.", vbExclamation + vbSystemModal, "Win Update Pause / Resume | by Abdullah ERTÜRK"
    End If
End If

Set objShell = Nothing