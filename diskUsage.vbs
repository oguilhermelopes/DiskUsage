strScomputer = InputBox("Enter the Server IP")
Const strReport = "diskUsage.txt"
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Set objLocator = CreateObject("WbemScripting.SWbemLocator") 
set objWMI = objLocator.ConnectServer(strScomputer, "root\cimv2")
Set colItems = objWMI.ExecQuery("Select * from Win32_LogicalDisk WHERE DriveType=3")

Dim strScomputer
Dim objWMIService, objItem, colItems, strInputFilename, FSO, WF, arrcomp
Dim strDriveType, strDiskSize, txt, writetextfile, arrComputer

Set FSO = CreateObject("Scripting.FileSystemObject") 
Set WriteTextFile = FSO.OpenTextFile(strReport, ForWriting, True)
WriteTextFile.WriteLine "Server" & vbtab & "Disk" & vbtab & "Disk Size" & vbtab & "Used Space" & vbtab & "Free Space"

For Each objItem in colItems
	DIM pctFreeSpace,strFreeSpace,strusedSpace
		strDiskSize = FormatNumber((objItem.Size/1073741824),2)
		strFreeSpace = FormatNumber((objItem.FreeSpace/1073741824),2)
		strUsedSpace = FormatNumber(((objItem.Size-objItem.FreeSpace)/1073741824),2)
WriteTextFile.WriteLine strScomputer & vbtab & objItem.Name & vbtab & strDiskSize & vbtab & strUsedSpace & vbTab & strFreeSpace
Next
msgbox ("DONE!!!")
