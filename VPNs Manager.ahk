/*
	VPNs Manager - v0.1
	Created by BLBC (github.com/hjk789)
	Copyright (c) 2020+ BLBC
*/

#SingleInstance force
SetWorkingDir %A_ScriptDir%
SetBatchLines 0		; For each line of code of the script, a predefined delay of 10 ms happens. This sets the delay to the smallest value needed for each line.

;*************************************
global isUsingComodoFirewall := false
global isUsingPeerBlock 	 := false

global peerblockVPNsFileFullPath 	:= "<FULL PATH TO WHERE IS LOCATED YOUR PEERBLOCK BLOCKLISTS>\VPNs.txt"
global firewallVPNsRegistryPath 	:= "HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\CmdAgent\CisConfigs\<CONFIG ID>\Firewall\Network Aliases\<THE NETWORK ZONE ID WHICH WILL CONTAIN THE VPNs>"

global vpnAdapterName 		 		:= "VPN Client Adapter - VPN"				; This is the default NIC name that SoftEther suggests for the network adapter created to be used by the VPNs. If you've set a different name, change it to to the name you've set.
global SoftEtherDirectoryPath		:= "C:\Program Files\SoftEther VPN Client"
global vpncmd 						 = "%SoftEtherDirectoryPath%\vpncmd.exe" 127.0.0.1 /client /cmd
;*************************************

RegRead, activeConfig, HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\CmdAgent\CisConfigs, Active		; This gets Comodo Firewall's currently active configuration ...
firewallVPNsRegistryPath := strReplace(firewallVPNsRegistryPath, "<CONFIG ID>", activeConfig)			; ... and adds it to the registry path.


global iniFileName := "vpns-settings.ini"
global whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
global chosenRow := ""
global highlightedRow := ""
global currentConName := ""
global currentConStatus := ""
global eligibleServers := []
global stopFetching


;/* Create the main screen */
;{
	gui, mainUI:new

	;/* Create the ListView */
	;{
		lvWidth := 460
		gui, add, listview, Sort -ReadOnly Grid h300 w%lvWidth% glistviewEvent, Name|Server|Good|Status|Last online|Chkd|Max speed|Max Average ;|Trend speed|Trend speed 2

		global cols := {"Name":1, "Server":2, "Good":3, "Status":4, "LastOnline":5, "Checked":6, "MaxSpeed":7, "MaxAverage":8}		; Store the columns indexes in an object for easy use.

		LV_ModifyCol(cols.server, "Integer Left")
	;}

	;/* Create the ImageList with the icons that will be used by the ListView */
	;{
		ImageListID := IL_Create()
		LV_SetImageList(ImageListID)

		IL_Add(ImageListID, "shell32.dll", 78) 		; Add the yellow triangle icon.
		IL_Add(ImageListID, "shell32.dll", 300)		; Add the green arrow icon.
		IL_Add(ImageListID, "shell32.dll", 132)		; Add the red X icon.
		IL_Add(ImageListID, "shell32.dll", 136)		; Add the globe with arrow icon.

		global icons := {"triangle":"Icon1", "greenArrow":"Icon2", "X":"Icon3", "globeWithArrow":"Icon4"}		; Store the icons indexes in an object for easy use.
	;}

	rebuildListView()

	;/* Create the buttons */
	;{
		Gui, Font, s22
		Gui, add, button, y15 w35 h35 gopenAddServerScreen, +

		Gui, Font, s30
		Gui, add, button, xp y+85 w34 h40 gchooseRandomServer, ↷
		Gui, add, button, xp y+85 w34 h40 grebuildListView, ⟲

		Gui, Font, s12
		Gui, add, button, x80 h40 gfetchServers, Fetch servers

		Gui, add, button, x+125 h40 gpingAllServers, Ping servers
	;}

	;/* Create the status bar */
	;{
		Gui, Font, s11
		Gui, Add, StatusBar
		statusbarWidth := lvWidth + 67
		SB_SetParts(statusbarWidth - 100)  			; Create a separation at the end of the status bar ...
		SB_SetText(LV_GetCount() " servers", 2)		; ... that displays the number of servers listed.
	;}

	;/* Create the context menu items */
	;{
		Menu, MainContextMenu, Add, Ping, contextMenuHandler
		Menu, MainContextMenu, Add	; Separator
		Menu, GoodSubContextMenu, Add, Route, contextMenuHandler
		Menu, GoodSubContextMenu, Add, Speed, contextMenuHandler
		Menu, GoodSubContextMenu, Add, Both, contextMenuHandler
		Menu, MainContextMenu, Add, Good, :GoodSubContextMenu
		Menu, MainContextMenu, Add
		Menu, MainContextMenu, Add, Copy, contextMenuHandler
		Menu, MainContextMenu, Add, Copy IP, contextMenuHandler
		Menu, MainContextMenu, Add
		Menu, MainContextMenu, Add, Delete, contextMenuHandler
	;}

	;/* Setup the tray options */
	;{
		Menu, Tray, NoStandard							; Remove all the pre-included items (Reload, Pause script, etc.).
		Menu, Tray, Add, Show VPNs Manager, ShowGui
		Menu, Tray, Add
		Menu, Tray, Add, Exit, mainUIGuiClose
		Menu, Tray, Default, Show VPNs Manager			; Make it so that when double-clicking the tray icon, the main screen is shown, as the "Show VPNs Manager" option is set as default.

		Menu, Tray, Click, 1							; Require only one click to trigger the default option.
		Menu, Tray, Icon, shell32.dll, 14				; Set the tray icon to the ringed globe icon.
		Menu, Tray, Tip, VPNs Manager					; Set the tray icon's tooltip.
	;}

	OnMessage(0x112, "WM_SYSCOMMAND")		; When minimized, run the WM_SYSCOMMAND function below, to hide the main GUI.


	;/* UI functions, subroutines and event handlers */
	;{
		;/*Minimize to tray and restore functions*/
		;{
			WM_SYSCOMMAND(wParam)	; This function is called whenever the Gui receives a command from the window's native controls, such as minimizing the window.
			{
			   If (wParam == 61472)  	; Minimize command.
				  SetTimer, HideGui, -1		; Hide the gui only after the window finishes being minimized, otherwise the minimize command makes it visible again. 1ms is enough time.
			}						   		; And it needs to be a timer so that the thread isn't interrupted in the meantime, otherwise the same thing happens.

			HideGui() {
			   Gui, mainUI:Hide
			}

			ShowGui() {
				Gui, mainUI:Show
			}
		;}

		mainUIGuiContextMenu()		; This function is called whenever the user right-clicks inside the mainUI GUI.
		{	
			highlightedRow := getRowData(A_EventInfo)		; When right-clicking a listview, the A_EventInfo variable value is the row index where the right-click occurred.
			Menu, MainContextMenu, Show
		}

		contextMenuHandler()	; This function is called, in a separate thread, whenever the user chooses an item of the context menu.
		{
			gui, mainUI:default		; LV_* functions use the default GUI of the thread it was called from. Each thread considers the most recently created GUI in it as the default one.
									; Because the contextMenuHandler is called from another thread, this command sets the mainUI GUI created in the main thread as the respective thread's default GUI.
			
			if (A_ThisMenu == "MainContextMenu")
			{
				if (A_ThisMenuItem == "Ping")
				{
					numSelectedRows := LV_GetCount("S")

					Loop %numSelectedRows%
					{
						RowNumber := LV_GetNext(RowNumber)
						if (!RowNumber)		; if 0
							break

						row := getRowData(RowNumber)
						
						try	
							success := pingServer(row.ip, RowNumber)
						catch 
							success := false
						
						if (!success)
						{
							msgbox Couldn't connect to the server.
							return
						}

						if (numSelectedRows > 1)
							SB_SetText("Pinged " A_Index " of " numSelectedRows " servers", 1)		; Show the progress in the status bar
					}

					sleep 2000

					SB_SetText("", 1)		; Clear the status bar after 2 seconds.
				}
				else if (A_ThisMenuItem == "Copy")
					clipboard := highlightedRow.name " " highlightedRow.server
				else if (A_ThisMenuItem == "Copy IP")
					clipboard := highlightedRow.ip
				else if (A_ThisMenuItem == "Delete")
				{
					chosenRow := highlightedRow
					deleteSelectedServers()
				}
			}
			else if (A_ThisMenu == "GoodSubContextMenu")
			{
				IniWrite, %A_ThisMenuItem%, %iniFileName%, % highlightedRow.ip, Good
				LV_Modify(highlightedRow.index, "Col" cols.good, A_ThisMenuItem)
			}
		}


		listviewEvent()		; This function is called whenever the listview receives an action, such as double-clicking a row.
		{
			if (A_GuiEvent == "DoubleClick")	; The A_GuiEvent variable value is the action that the listview received.
				connectToServer(A_EventInfo)	; Here A_EventInfo value is the row index where the action was performed.
		}
	;}

	gui show,, VPNs Manager
;}


;/* Startup parameters */
;{
	if (A_Args.MaxIndex() >= 1)
	{
		Critical		; Prevent the startup procedures from being interrupted.

		if (A_Args[1] == "--fetchServers")
			fetchServers()
		else if (A_Args[1] == "--connectOnStartup")
		{
			chooseRandomServer()

			;/* When the startup argument connectOnStartup is used, update the servers status once after the first 3 minutes, which is enough time for everything to be up and running before that. */
			tempFunc := Func("pingAllServers").bind()		; Because there can be only one timer for each function/label, this creates a function object and binds an empty parameter to make it seem different to the interpreter.
			setTimer, %tempFunc%, -120000		; 2 minutes.
		}
		else if (A_Args[1] == "--openAddServerScreen")
			openAddServerScreen()
		else if (inStr(A_Args[1], "--delete"))
		{
			deleteSelectedServers(A_Args[2], A_Args[3])

			if (A_Args[1] == "--deleteBadServer")
				chooseRandomServer()
		}

		Critical off
	}
;}


setTimer, pingAllServers, 3600000		; Ping all servers each hour.

setTimer, waitForConnectError, -1		; Wait, in a separated thread, for a connection error, otherwise it would halt the whole program.


global avgarr := []
global avgsamples := 50
global trendobj := {}
global trend := "0"
global trend2 := "0"
global maxspeed := 0
global maxavg := 0

setupNetworkSpeedMeter()

return




rebuildListView()
{
	gui, mainUI:default

	LV_Delete()

	IniRead, ips, %iniFileName%			; Get a list with the name of all sections in the ini file, which are all IP addresses, plus the Settings section.
	getCurrentConnectionFromCmd()


	Loop, Parse, ips, `n				; Loop line by line.
	{
		icon := "Icon0"					; Default icon (none).

		if (A_LoopField == "Settings")
			continue					; Ignore the Settings section.

		;/* Get the servers data from the ini file */
		;{
			IniRead, name, %iniFileName%, %A_LoopField%, Name
			IniRead, port, %iniFileName%, %A_LoopField%, Port
			IniRead, good, %iniFileName%, %A_LoopField%, Good, %A_Space%
			IniRead, status, %iniFileName%, %A_LoopField%, Status
			IniRead, lastSeenOnline, %iniFileName%, %A_LoopField%, LastSeenOnline, %A_Space%
			IniRead, checked, %iniFileName%, %A_LoopField%, Checked
		;}

		;/* Determine the icons to be shown in the listview */
		;{
			if (name == currentConName)
			{
				if (inStr(currentConStatus, "|Connection Completed"))
					icon := icons.greenArrow
				else if (inStr(currentConStatus, "not connected"))
					icon := icons.triangle
				else
					icon := icons.globeWithArrow
			}
			
			if (status == "Off")
				icon := icons.X
		;}

		LV_Add(icon, name, A_LoopField ":" port, good, status, lastSeenOnline, checked)
	}

	LV_ModifyCol()
	LV_ModifyCol(cols.name, "120")
	LV_ModifyCol(cols.good, "SortDesc AutoHdr")
	LV_ModifyCol(cols.status, "SortDesc")
	LV_ModifyCol(cols.checked, "37")

	buildListEligibleServers()

	SB_SetText(LV_GetCount() " servers", 2)
}


getCurrentConnectionFromCmd()
{
	FileRead, currentConStatus, %A_ScriptFullPath%:currentConStatus		; Here, and in other parts, NTFS' alternate data streams are used as a disposable, persistent storage of vpncmd's outputs, instead of creating files just for that.

	runwait, cmd /c (%vpncmd% accountlist) > %iniFileName%:serversList,, hide
	FileRead, serversList, %iniFileName%:serversList
	RegExMatch(serversList, "Name \|([^\n]+?)\nStatus +?\|(?:Connected|Connecting).+?(\d+\.\d+\.\d+\.\d+)", currentConMatch)		; Parse vpncmd's output to get the name of the connected server. The parsed output looks like this:		VPN Connection Setting Name |USA, New York
																																																									;		Status                      |Connected
	currentConName := currentConMatch1		; RegExMatch creates a pseudo array containing the matches. In this case, the currentConMatch variable contains	the
											; whole matched string, and currentConMatch1 contains the first capturing group in the regex, which is the server's name.
	
    if (currentConName != "")
	{
		if (chosenRow == "")
		{
			chosenRow := {}
			chosenRow.name := currentConName
			chosenRow.ip := currentConMatch2
		}
		
		checkConnection()
	}
}


buildListEligibleServers()
{
	Critical 
	
	gui, mainUI:default

	eligibleServers := []

	loop % LV_GetCount()
	{
		row := getRowData(A_Index)

		if (row.status == "Online" && (row.checked == "no" || row.good != ""))
			eligibleServers.push(A_Index)

		if (currentConName == row.name)
			chosenRow := row
	}
	
	Critical off
}


getRowData(index)
{
	gui, mainUI:default

	LV_GetText(name, index, cols.name)
	LV_GetText(server, index, cols.server)
	ip := (strSplit(server, ":"))[1]
	LV_GetText(good, index, cols.good)
	LV_GetText(status, index, cols.status)
	LV_GetText(checked, index, cols.checked)

	return {"index": index, "name": name, "server": server, "ip": ip, "good": good, "status": status, "checked": checked}
}


waitForConnectError()
{
	loop
	{
		winwaitactive Connect Error		; Wait for the connect error dialog when the "Hide Status Window" option is disabled. The alternative to this would be to keep
		winhide							; polling for the connection status, but it would need to have a short interval like 1 second, and that may not be a good solution.
		chooseRandomServer()
	}
}


chooseRandomServer()
{
	random, rand, 1, eligibleServers.MaxIndex()

	connectToServer(eligibleServers[rand])

	eligibleServers.removeAt(rand)		; Delete it from the array so that it's not chosen again until the next time the eligible servers list is built.

	if (!eligibleServers.MaxIndex())	; When all eligible servers were already chosen,
		buildListEligibleServers()		; rebuild the list.	
}


connectToServer(index)
{
	gui, mainUI:default

	runwait %vpncmd% accountdisconnect "%currentConName%",, hide		; SoftEther requires the current connection to be disconnected before connecting to another server.
	if (chosenRow == "")
		rebuildListView()
	else
		LV_Modify(chosenRow.index, "Icon99")		; Invalid icon index to remove the current icon.

	chosenRow := getRowData(index)
	name := chosenRow.name

	run %vpncmd% accountconnect "%name%",, hide

	winwait Connecting,, 2		; Wait for the connecting dialog when the "Hide Status Window" option is disabled.
	ifwinexist Connecting		; In case that option is enabled, the 2 seconds timeout prevents the script from getting stuck infinitely.
		winhide

	replaceFileContent(A_ScriptFullPath ":currentConStatus", "")
	currentConStatus := ""

	currentConName = %name%


	LV_Modify(index, icons.globeWithArrow)

	setTimer, checkConnection, 1000
}


replaceFileContent(filePath, fileContent)
{
	file := FileOpen(filePath, "w")		; Create an empty file. If the file already exists, overwrite it.
	file.Write(fileContent)
	file.Close()
}


checkConnection()
{
	gui, mainUI:default

	name := chosenRow.name
	ip := chosenRow.ip

	runwait, cmd /c (%vpncmd% accountStatusGet "%name%") > "%A_ScriptFullPath%:currentConStatus",, hide

	FileRead, currentConStatus, %A_ScriptFullPath%:currentConStatus
	IniRead, checked, %iniFileName%, %ip%, Checked

	if (inStr(currentConStatus, "not connected") || inStr(currentConStatus, "Retrying") || (inStr(currentConStatus, "|Connection Completed") && inStr(currentConStatus, "MD5")))
	{
		if (inStr(currentConStatus, "MD5"))
			msgbox This server uses RC4-MD5 encryption and will be deleted.
		else if (checked != "yes")
			msgbox, 49,, Couldn't connect to the server. Do you want to delete this server?
		else
			msgbox, 53,, Couldn't connect to the server.

		setTimer, checkConnection, delete
		replaceFileContent(A_ScriptFullPath ":currentConStatus", "")
		currentConStatus := ""

		ifmsgbox Cancel
			return
			
		ifmsgbox Retry
		{
			connectToServer(chosenRow.index)
			return
		}			

		if (checked != "yes")
		{
			if (!A_IsAdmin)
				Run, *RunAs "%A_ScriptFullPath%" --deleteBadServer %name% %ip%		; Start the VPNs Manager again as admin, and because the SingleInstance directive is set to "force", the unelevated instance is automatically killed.

			deleteSelectedServers(name, ip)
		}
	}
	else if (inStr(currentConStatus, "|Connection Completed"))
	{
		setTimer, checkConnection, delete

		if (checked != "yes")
			IniWrite, yes, %iniFileName%, %ip%, Checked

		if (chosenRow != "")
			LV_Modify(chosenRow.index, icons.greenArrow " Col" cols.checked, "yes")
	}
}


deleteSelectedServers(pName := "", pIP := "")
{
	gui, mainUI:default

	if (!A_IsAdmin)
	{
		msgbox, 49,, Are you sure you want to delete the selected server(s)?
		ifmsgbox Cancel
			return
	
		index := chosenRow.index
		name := chosenRow.name
		ip := chosenRow.ip
		Run, *RunAs "%A_ScriptFullPath%" --delete "%name%" %ip%
	}
	else if (pName != "")
	{
		msgbox, 49,, Are you sure you want to delete the server "%pName%"?
		ifmsgbox Cancel
			return
	}


	RowNumber = 0

	IniRead, deletedServers, %iniFileName%, Settings, DeletedServers, %A_Space%


	Critical		; Prevent the user from deselecting the rows before it finished deleting.
	
	numSelectedRows := LV_GetCount("S")

	if (numSelectedRows == 0)	
		numSelectedRows = 1
	
	Loop %numSelectedRows%
	{
		if (pName == "")
		{
			RowNumber := LV_GetNext(RowNumber - 1)
			if (!RowNumber)
				break

			row := getRowData(RowNumber)
			ip := row.ip
			name := row.name

			SB_SetText(LV_GetCount() " servers", 2)
		}
		
		;/* Remove from SoftEther VPNs list*/
		runwait, %vpncmd% accountdelete "%name%",, hide
		
		deletedServers .= ip ","

		;/* Remove the server from the ini file */
		IniDelete, %iniFileName%, %ip%
		
		;/* Remove from the PeerBlock list */
		if (isUsingPeerBlock)
		{
			FileRead, peerblockVPNsFileContent, %peerblockVPNsFileFullPath%
			peerblockVPNsFileContent := strReplace(peerblockVPNsFileContent, name ":" ip "-" ip, "")
			replaceFileContent(peerblockVPNsFileFullPath, peerblockVPNsFileContent)
		}

		;/* Remove from the firewall's exceptions */
		if (isUsingComodoFirewall)
		{
			loop, reg, %firewallVPNsRegistryPath%, k
			{
				RegRead, regIP, %firewallVPNsRegistryPath%\%A_LoopRegName%\IPV4, AddrStart
				if (regIP == ip)
				{
					RegDelete %firewallVPNsRegistryPath%\%A_LoopRegName%
					break
				}
			}
		}

		if (RowNumber > 0)
			LV_Delete(RowNumber)
		else break
	}
	
	Critical off

	IniWrite, %deletedServers%, %iniFileName%, Settings, DeletedServers

	if (RowNumber == 0)
		rebuildListView()
}


pingAllServers()
{
	gui, mainUI:default

	numRows := LV_GetCount()

	Loop, %numRows%
	{
		row := getRowData(A_index)

		try	
			success := pingServer(row.ip, A_Index)
		catch 
			success := false
		
		if (!success)
		{
			msgbox, 53,, Couldn't connect to the server.

			ifmsgbox Retry
			{
				pingAllServers()
				return
			}
			ifmsgbox Cancel
				return
		}

		SB_SetText("Pinged " A_Index " of " numRows " servers", 1)
	}

	rebuildListView()

	sleep 2000

	SB_SetText("", 1)
}


pingServer(serverIp, rowIndex)
{
	gui, mainUI:default

	ip := serverIp
	IniRead, port, %iniFileName%, %ip%, Port

	Critical		; Enable the critical mode for the request, otherwise the request fails when another thread interrupts it.
	
	whr.open("POST", "https://ports.yougetsignal.com/check-port.php")			; The port check feature from YouGetSignal is used to check if the servers port used for the VPN is open. The alternatives to this would be using a third-party commandline 
	whr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")	; program that enables checking ports or using the PowerShell cmdlet Test-Connection. Using a third-party server is the most compatible and standalone method.
	whr.send("remoteAddress=" ip "&portNumber=" port)

	response := whr.responseText
	success := inStr(response, ip)		; The server's response always include the pinged IP, either when open or closed. This checks whether the request went OK and whether the response includes the expected info.
	
	Critical off

	if (success)
	{
		formatTime, now,, dd/MM HH:mm		; Get the current time.
		isOpen := inStr(response, "open")

		if (isOpen)
		{
			status := "Online"
			IniWrite, %now%, %iniFileName%, %ip%, LastSeenOnline
			LV_Modify(rowIndex, "Col" cols.lastonline, now)
		}
		else status := "Off"

		IniWrite, %status%, %iniFileName%, %ip%, Status
		LV_Modify(rowIndex, "Col" cols.status, status)
	}
	else if (inStr(response, "check limit reached"))		; YouGetSignal has a daily limit of port checks. But you only reach this limit if you repeatedly ping several servers several times. 
	{														; Pinging 20 servers 10 times a day isn't enough. Also, the limit is per IP, so you just need to switch to another VPN server to use it again.
		msgbox % (strSplit(response, ". "))[1]				; When the limit is reached, YouGetSignal responds with an error message. This shows in a msgbox the most relevant part of the error.
	}
	
	
	return success
}


addNewServer(serverName, serverIP, port)
{
	gui, mainUI:default

	;/* Add the server to SoftEther VPNs list */
	runwait %vpncmd% accountcreate "%serverName%" /server:%serverIP%:%port% /hub:"VPNGATE" /username:"vpn" /nic:"%VpnAdapterName%",, hide

	;/* Add to the ini file */
	IniWrite, %port%, %iniFileName%, %serverIP%, Port
	IniWrite, %serverName%, %iniFileName%, %serverIP%, Name
	IniWrite, Online, %iniFileName%, %serverIP%, Status
	IniWrite, no, %iniFileName%, %serverIP%, Checked

	;/* Add to Comodo Firewall's exceptions */
	if (isUsingComodoFirewall)
	{
		Random, rand, 50, 999
		RegWrite, REG_DWORD, %firewallVPNsRegistryPath%, Num, 999
		RegWrite, REG_DWORD, %firewallVPNsRegistryPath%\%rand%, Source, 2
		RegWrite, REG_DWORD, %firewallVPNsRegistryPath%\%rand%, Type, 1
		RegWrite, REG_SZ, %firewallVPNsRegistryPath%\%rand%\IPV4, AddrStart, %serverIP%
		RegWrite, REG_SZ, %firewallVPNsRegistryPath%\%rand%\IPV4, AddrEnd, %serverIP%
		RegWrite, REG_DWORD, %firewallVPNsRegistryPath%\%rand%\IPV4, AddrType, 1
	}

	;/* Add to the PeerBlock list */
	if (isUsingPeerBlock)
		FileAppend, %serverName%:%serverIP%-%serverIP%`n, %peerblockVPNsFileFullPath%
}


changeSoftEtherConfig()
{
	gui, mainUI:default

	FileRead, softetherConfigFileContent, %SoftEtherDirectoryPath%\vpn_client.config
	softetherConfigFileContent := strReplace(softetherConfigFileContent, "DisableQoS false"       , "DisableQoS true")
	softetherConfigFileContent := strReplace(softetherConfigFileContent, "NoUdpAcceleration false", "NoUdpAcceleration true")
	softetherConfigFileContent := strReplace(softetherConfigFileContent, "UseCompress false"      , "UseCompress true")
	softetherConfigFileContent := strReplace(softetherConfigFileContent, "HideNicInfoWindow false", "HideNicInfoWindow true")
	softetherConfigFileContent := strReplace(softetherConfigFileContent, "NumRetry 4294967295"    , "NumRetry 1")
	softetherConfigFileContent := strReplace(softetherConfigFileContent, "RetryInterval 15"       , "RetryInterval 1")

	runwait, net stop sevpnclient,, hide		; SoftEther's service need to be stopped to be able to change the config file.

	replaceFileContent(SoftEtherDirectoryPath "\vpn_client.config", softetherConfigFileContent)

	runwait, net start sevpnclient,, hide

	run, %SoftEtherDirectoryPath%\vpncmgr.exe

	if (currentConName != "")
		run %vpncmd% accountconnect "%currentConName%",, hide
}


fetchServers()
{
	if (!A_IsAdmin)
		Run *RunAs "%A_ScriptFullPath%" --fetchServers

	;/* Create the "Fetching servers..." dialog */
	;{
		gui fetchingUI:new, -Caption Border
		gui, font, s12
		gui add, text,, Fetching servers...
		gui add, button, xp+37 y+10 gSetStopFetching, Stop
		gui fetchingUI:show		
		stopFetching := false
	;}

	gui, mainUI:default

	addedServers := 0
	acceptableServerRegex := "s)<h3 class=["" -\w""]+?>(\d+\.\d+\.\d+\.\d+)<\/h3>\s+<p class=["" -\w""]+?>\w+ TCP\((\d+)\)[ \w()]+?<br>ping (80|[1-7][0-9])ms\(US\)"  ; Regex to search for a server with a latency of 10-80ms to US, capturing its IP and TCP port.

	loop																						; Fetch infinitely ...
	{
		if (addedServers >= 15 || stopFetching)													; ... until 15 servers are added or until the user clicks the "Stop" button.
		{
			gui fetchingUI:destroy
			break
		}

		try 
		{
			whr.open("GET", "https://freevpn.gg?p=" A_Index)										; VPNs Manager fetches the servers from freevpn.gg, which holds a database of VPN servers, mostly, if not all, from VPN Gate. The 
			whr.send()																				; advantage over taking directly from VPN Gate is that the latency tests are made from a US server instead of from a Japanese one, 
			pageResponse := whr.responseText														; which lets you have a more precise idea of the server's latency. Also, because it stores the full list of servers, you can even 
																									; filter by country if you want, just add "/s/two_letters_country_code" before the question mark, e.g. "https://freevpn.gg/s/US?p=".
		}
		catch 
		{
			msgbox, 48,, Couldn't connect to the server.
			if (addedServers > 0)
				break
			else
				return
		}
		
		currentPage := A_Index


		i := 1  ; Position at the page's HTML code
		while (i := RegExMatch(pageResponse, acceptableServerRegex, match, i+StrLen(match)))	; RegExMatch returns the character index where the occurrence begins. If it's not found, it returns 0, which is interpreted as false and the
		{																						; while loop ends. The last parameter determines where the search must start from, which in this case is from where the previous occurrence ends.
			
			FileRead, vpnsIniFileContent, %iniFileName%
																								
			ip := match1  																		; match1 contains the first capturing group in the regex, which is the IP, 
			port := match2																		; and match2 contains the second group, which is the port number.
			country := city := ""
				
			if (ip != "" && !inStr(vpnsIniFileContent, ip))	  									; Prevent an already added or deleted server from being added again. Only proceed if this IP doesn't appear anywhere in the file.
			{												   	  								
				whr.open("GET", "https://freevpn.gg/c/" ip)										; Open the fetched server's details page ...
				whr.send()
				locationResponse := whr.responseText
				RegExMatch(locationResponse, "<br><br>(.+?)<br>", location)						; ... to parse the server's location, which generally is "City, Country". Some servers have only the country name.

				if (inStr(location1, ","))						 								; If it's "City, Country" ...
				{
					locationsplit := StrSplit(location1, ", ")   								; ... separate the two to do the processings.
					country := locationsplit[2]
					city := locationsplit[1]
				}
				else																			; else, if it's only the country name.
					country := location1

				;/* Replace the countries name with their short versions to save space in the listview and make it easier to read */
				country := strReplace(country, "United States of America", "USA")
				country := strReplace(country, "Russian Federation", "Russia")
				country := strReplace(country, "United Kingdom of Great Britain and Northern Ireland", "UK")
				country := strReplace(country, "Venezuela (Bolivarian Republic of)", "Venezuela")
				country := strReplace(country, "Republic of ", "")

				StrReplace(vpnsIniFileContent, country, country, count)							; Get how many servers of this country are already added. The fourth parameter outputs the number of occurrences 
																								; replaced. As all occurrences of the country name were replaced by itself, the result is basically a counter.
				location := country
				if (city != "")
					location .= ", " city														; Here the location name is reversed, in which the most relevant info, the country 
                                                                                                ; name, comes before the city name. The ".=" operator is shorthand for concatenation.

				if (count < 5)																	; Prevent adding more than 5 servers of the same country, for variety reasons. Otherwise it would fetch only 
				{                                                                               ; servers from Korea and Japan, as the vast majority of VPN Gate servers are from these two countries.
				
					if (inStr(vpnsIniFileContent, "Name=" location))							; If there's already a server with this name ...
					{
						random, rand, 1, 50
						location .= " " rand													; ... append a random number to it to differentiate.
					}


					addNewServer(location, ip, port)


					addedServers++

					LV_Add("Icon0", location, ip)

					SB_SetText("Page " currentPage ", " addedServers " added - " location, 1)
				}
			}
			
		}
		
		SB_SetText("Page " currentPage ", " addedServers " added - " location, 1)
	}


	finishAddServer()

	sleep 2000
	SB_SetText("", 1)
}

setStopFetching() {
	Critical
	stopFetching := true
}


finishAddServer()
{
	changeSoftEtherConfig()

	if (isUsingComodoFirewall)
	{
		run "C:\Program Files\COMODO\COMODO Internet Security\cis.exe" --configUI=CeFirewallSettingsPage.html		; Open Comodo Firewall's settings screen ...
		winwaitactive ahk_class CisMainWizard
		sleep 500
		send {tab}{enter}                                                                                           ; ... and automatically confirm it. This is necessary for the changes in the registry to take effect.
	}

	if (isUsingPeerBlock)
		msgbox Don't forget to update PeerBlock settings.
	
	gui addServerUI:destroy

	rebuildListView()	
}


openAddServerScreen()
{
	if (!A_IsAdmin)
		Run, *RunAs "%A_ScriptFullPath%" --openAddServerScreen

	gui, addServerUI:new

	gui, font, s10

	;/* Name field */
	gui, add, text, x18 y8, Name
	gui, add, edit, vnameInput x18 y30 w208 limit60


	;/* IP field */
	gui, add, text, x18 y60	, Server IP
	gui, add, edit, vipInput x18 y80 limit15

	gui, font, s15
	gui, add, text, x171 y78, :

	;/* Port field */
	gui, font, s10
	gui, add, text, x180 y60, Port
	gui, add, edit, vportInput w45 x180 y80 limit5 number

	
	;/* "Add" button */
	Gui, Font, s12
	gui, add, button, x23 y+15 w80 h30 default gSubmitAddServerScreen, Add
		
	;/* "Finish" button */
	gui, add, button, xp+120 yp w80 h30 gFinishAddServer, Finish
	

	gui show, w245 h160, Add Server
}

submitAddServerScreen()
{
	global nameInput, ipInput, portInput	; The input box variables are required to be global.

	Gui, Submit, NoHide

	addNewServer(nameInput, ipInput, portInput)
	
	ControlSetText, Edit1, 		; Clear the input boxes when done adding.
	ControlSetText, Edit2, 
	ControlSetText, Edit3, 
}


;/* Network speed meter functions */
;{
	setupNetworkSpeedMeter()
	{	
		;/* The code inside this function was created by Sean, with few adaptations by BLBC. Original code: https://autohotkey.com/board/topic/16574-network-downloadupload-meter/  */

		If GetInterfaceTable(tb)
			return

		Loop, % DecodeInteger(&tb)
		{
			If DecodeInteger(&tb + 4 + 860 * (A_Index - 1) + 544) < 4 || DecodeInteger(&tb + 4 + 860 * (A_Index - 1) + 516) = 24
				Continue
			ptr := &tb + 4 + 860 * (A_Index - 1)
				Break
		}

		If !ptr
			return

		SetTimer, updateNetworkMeter, 1000

		settimer, resetIdle, 60000
	}

	updateNetworkMeter()
	{
		num := getDownloadedKilobytes()

		if (avgarr.maxindex() >= avgsamples)
			avgarr.removeAt(1)
			
		avgarr.push(num)

		if (num > maxspeed)
		{
			maxspeed := floor(num)
			LV_Modify(chosenRow.index, "Col" cols.MaxSpeed, maxspeed)
		}

		avgres := 0

		loop, % avgarr.maxindex()
			avgres += avgarr[A_index]

		avgres := avgres / avgsamples

		if (avgres > maxavg)
		{
			maxavg := floor(avgres)
			LV_Modify(chosenRow.index, "Col" cols.MaxAverage, maxavg)
		}
		
	}

	getDownloadedKilobytes()
	{
		;/* The code inside this function was created by Sean, with adaptations by BLBC. Original code: https://autohotkey.com/board/topic/16574-network-downloadupload-meter/  */

		global

		DllCall("iphlpapi\GetIfEntry", "Uint", ptr)

		totalDownloadedBytes := DecodeInteger(ptr + 552)

		downloadedKilobytes := (totalDownloadedBytes - previousTotalDownloadedBytes) / 1024

		previousTotalDownloadedBytes := totalDownloadedBytes

		return downloadedKilobytes
	}


	;/* The four functions below were created by Sean (https://autohotkey.com/board/topic/16574-network-downloadupload-meter/). */

	DecodeInteger(ptr)
	{
		Return *ptr | *++ptr << 8 | *++ptr << 16 | *++ptr << 24
	}

	GetInterfaceTable(ByRef tb, bOrder = False)
	{
		nSize := 4 + 860 * GetNumberOfInterfaces() + 8
		VarSetCapacity(tb, nSize)
		Return DllCall("iphlpapi\GetIfTable", "Uint", &tb, "UintP", nSize, "int", bOrder)
	}

	GetNumberOfInterfaces()
	{
		DllCall("iphlpapi\GetNumberOfInterfaces", "UintP", nIf)
		Return nIf
	}

	GetIfEntry(ByRef tb, idx)
	{
		VarSetCapacity(tb, 860)
		DllCall("ntdll\RtlFillMemoryUlong", "Uint", &tb + 512, "Uint", 4, "Uint", idx)
		Return DllCall("iphlpapi\GetIfEntry", "Uint", &tb)
	}
;}


return
mainUIGuiEscape:
mainUIGuiClose:
exitapp