' pInven.vbs

' Script for asset inventory following Andrei guidelines 
' Dump results to CSV file (hostname.csv)
' Version: 0.0.4

' 02/15/2013 - Sergio Acosta

On Error Resume Next

Dim oLogObject
Dim oLogOutput
Dim sLogFile

set oLogOutput = nothing

Dim aArgs
Dim strComputer
Dim objWMIService
Dim colItems
Dim objItem

'If there are no args
If Wscript.Arguments.Count = 0 Then
   strComputer = "localhost"
Else
   strComputer = Wscript.Arguments(0)
End If

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
'Exit codes
if err Then
  if err.number = 462 Then
    Wscript.Echo "Error: Impossible to find  machine \\" & strComputer
    Wscript.Quit 2
  else
    WScript.Echo "Fatal Error: " & err.number & " - " & err.description 'Error fatal!
    Wscript.Quit 3
  end if    
end if

sLogFile = strComputer & ".csv"

set oLogObject = CreateObject("Scripting.FileSystemObject")

'Add a new line to CSV
If oLogObject.FileExists(sLogFile) Then
	set oLogOutput = oLogObject.OpenTextFile (sLogFile, 8, True)
Else
	set oLogOutput = oLogObject.CreateTextFile(sLogFile)
		oLogOutput.WriteLine "ASSET,ASSET TYPE,PURCHASE DATE,COMPANY,SITE,DEPARTMENT,COMPUTER NAME, USER, PRODUCT ID, SERIAL,BRAND,MODEL,MEMORY,PROCESSOR,OS"
End If

if err Then
  if err.number = 462 Then
    Wscript.Echo "Error: Could not find a machine named " & strComputer
    Wscript.Quit 2
  else
    WScript.Echo "Error: " & err.number & " - " & err.description
    Wscript.Quit 3
  end if    
end if

'Serial number & part number
Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_BaseBoard",,48) 

For Each objItem in colItems 
	SerialNumber = objItem.SerialNumber
	PartNumber = objItem.PartNumber
Next

'OS
Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_OperatingSystem",,48) 

For Each objItem in colItems 
	OS = objItem.Caption
Next

'Processor
Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_Processor",,48) 

For Each objItem in colItems 
	Processor = objItem.Name
Next

'Memory 

Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_ComputerSystem",,48) 

For Each objItem in colItems 
     
    ramMemory = objItem.TotalPhysicalMemory 

Next

PartNumberReg = ReadReg("HKLM\HARDWARE\DESCRIPTION\System\BIOS\SystemSKU")

'Check if PartNumber is OK (Not in use from 2/18/2012)
'If Len(PartNumber) < 3  Then
'PartNumber = "N/A"
'End If

'More useful data
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_ComputerSystem",,48) 

For Each objItem in colItems 

      oLogOutput.WriteLine ",,,,,," & objItem.Name & "," & _
      """" & RemoveDomain(objItem.UserName) & """," & _
	  PartNumberReg & "," & _
      SerialNumber & "," & _
      objItem.Manufacturer & "," & _
      """" &  objItem.Model & """," & _
      """" & ConvertSize(ramMemory) & """," & _
      """" & Processor & """," & _
	  """" & OS & """," & _
      """" & PartNumber & """"
Next

oLogOutput.Close


'Dim FSO 
'Set FSO = CreateObject("Scripting.FileSystemObject")
'strFile = ""
'strRename = ""
'If FSO.FileExists(strFile) Then
	'FSO.MoveFile strFile, strRename
'End If

'
Msgbox "Finished."
wscript.quit (0)


'Convert RAM to MB, GB, TB... 
Function ConvertSize(Size) 
Do While InStr(Size,",") 
    CommaLocate = InStr(Size,",") 
    Size = Mid(Size,1,CommaLocate - 1) & _ 
        Mid(Size,CommaLocate + 1,Len(Size) - CommaLocate) 
Loop

Suffix = " Bytes" 
If Size >= 1024 Then suffix = " KB" 
If Size >= 1048576 Then suffix = " MB" 
If Size >= 1073741824 Then suffix = " GB" 
If Size >= 1099511627776 Then suffix = " TB" 

Select Case Suffix 
    Case " KB" Size = Round(Size / 1024, 1) 
    Case " MB" Size = Round(Size / 1048576, 1) 
    Case " GB" Size = Round(Size / 1073741824, 1) 
    Case " TB" Size = Round(Size / 1099511627776, 1) 
End Select

ConvertSize = Size & Suffix 
End Function

'Remove domain from username if existes
Function RemoveDomain(userName)
	name = Split(Username,"\")
	RemoveDomain = Ucase(name(1)) 'Convert domain to uppercase
End Function

'Simply registry read function
Function ReadReg(RegPath)
On Error Resume Next
     Dim objRegistry, Key
     Set objRegistry = CreateObject("Wscript.shell")
     Key = objRegistry.RegRead(RegPath)
     ReadReg = Key
End Function
