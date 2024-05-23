
' INFO:

' Windows Update Pause / Resume
' Copyright (C) 2024 Abdullah ERT�RK
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
    MsgBox "Windows G�ncelle�tirme Duraklatma S�resi varsay�lan ayara (5 hafta) d�n��t�r�ld�." & vbCR & "" & vbCR & "Windows Update hizmeti etkinle�tirildi.", vbInformation + vbSystemModal, "Win Update Pause / Resume | by Abdullah ERT�RK"

	Set objShell = CreateObject("WScript.Shell")
	strPowerShellCmd1 = "Stop-Process -Name SystemSettings -Force"
	strPowerShellCmd2 = "Start-Process ms-settings:windowsupdate"
	objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd1 & """", 0, True
	objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd2 & """", 0, True

	Set objShell = Nothing

Else

    PauseWeeks = InputBox("Windows G�ncelle�tirmeleri Duraklat se�ene�ine eklemek istedi�iniz hafta say�s�n� girin:" & vbCR & "" & vbCR & "�rnek: 10, 20, 30, 40, vb.", "G�ncelle�tirmeleri Duraklatma S�resi")

    If IsNumeric(PauseWeeks) Then
        PauseWeeks = CInt(PauseWeeks)
        If PauseWeeks > 0 Then
            FlightSettingsMaxPauseDays = PauseWeeks * 7 

            objShell.RegWrite "HKEY_LOCAL_MACHINE\" & strKeyPath & "\FlightSettingsMaxPauseDays", CInt(FlightSettingsMaxPauseDays), "REG_DWORD"

            objShell.Run "cmd /c net start wuauserv", 0, True

            MsgBox "Windows G�ncelle�tirmeleri Duraklat se�ene�ine " & PauseWeeks & " hafta (" & FlightSettingsMaxPauseDays & " g�n) eklendi.", vbInformation + vbSystemModal, "Win Update Pause / Resume | by Abdullah ERT�RK"
	
			Set objShell = CreateObject("WScript.Shell")
			strPowerShellCmd1 = "Stop-Process -Name SystemSettings -Force"
			strPowerShellCmd2 = "Start-Process ms-settings:windowsupdate"
			objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd1 & """", 0, True
			objShell.Run "powershell.exe -ExecutionPolicy Bypass -Command """ & strPowerShellCmd2 & """", 0, True
			
			Set objShell = Nothing

        End If
    Else
        MsgBox "L�tfen ge�erli bir say� girin.", vbExclamation + vbSystemModal, "Win Update Pause / Resume | by Abdullah ERT�RK"
    End If
End If

Set objShell = Nothing