REM  *****  BASIC  *****
option explicit
option base 0
' 
'   Licensed under the Apache License, Version 2.0 (the "License"); you may not
'   use this file except in compliance with the License. You may obtain a copy of
'   the License at
'  
'     http://www.apache.org/licenses/LICENSE-2.0
'  
'   Unless required by applicable law or agreed to in writing, software 
'   distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
'   WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
'   License for the specific language governing permissions and limitations under
'   the License. 
'

Sub Main

End Sub


global oDialog1, oDialog2
global sUsername, sPassword

' -1=don't use, 0=not probed yet (default), 1=in use
global iUsage

global iNumTargets
global oTargets(100), bTargetsActive(100)
global sCurrSels(100), sPrevSels(100), lSelTimers(100)
global oCellListeners(100), oSelListeners(100)
global bIsWatchings(100), bStopWatchings(100)
global oChatWindow, sChatUrl

global oConnections(100)
global sConnectionUrls(100)
global bConnectionUsed(100)
global iConnectionUseCount(100)

' (per se false, don't start it twice)
global bLoopRunning

' more configuration
global iWaitTime

' -------- MAIN ENTRY POINTS --------

' Method to hook into document open (INCOMPLETE! - discard?!)
Sub Rangereplication_Init_Background(byval oEvent)
	if instr(oEvent.Source.dbg_properties, "ScModelObj")>0 then
		if oEvent.Source.Sheets.hasByName("ggglobal_shadow") then
			' if user says "don't" then let go
			if iUsage=0 then
				perform_logon
			end if
			if iUsage=1 then
				' ensure cellupd in place
				install_cellupd
				' upstream and downstream listening
				'TODO: for all that have shadow sheets
				assign_listeners(oEvent.Source)
				listen_for_changes(oEvent.Source)
			end if
		end if
	end if
end sub	

' Method to hook into document close (INCOMPLETE! - discard?!)
Sub Rangereplication_Shutdown_Background(byval oEvent)
	' stop listening and unassign listeners
	'TODO: for all that have shadow sheets
	stop_listen_for_changes
	unassign_listeners(oEvent.Source)
end sub

' Method to hook into the menu
Sub Rangereplication_Start_Listening
	if thisComponent.Sheets.hasByName("ggglobal_shadow") then
		' allow several attempts for login
		if iUsage<>1 then
			perform_logon
		end if
		if iUsage=1 then
			' check for DB availability (and gracefully exit if not)
			if check_db_connection()=false then
				exit sub
			end if
			' ensure cellupd in place
			install_cellupd
			' start when not yet listening
			assign_listeners(thisComponent)
			listen_for_changes(thisComponent)
		end if
	else
		' offer to show the edit settings dialog
		if msgbox("No database specified yet. Specify now?", 4)=6 then
			Rangereplication_Edit_Shadow
			' then offer to re-try
			if thisComponent.Sheets.hasByName("ggglobal_shadow") then
				if msgbox("Retry start now?", 4)=6 then
					Rangereplication_Start_Listening
				end if
			end if
		end if
	end if
end sub

' Method to hook into the menu
Sub Rangereplication_Stop_Listening
	' stop listening and unassign listeners
	stop_listen_for_changes
	unassign_listeners(thisComponent)
end sub

' Method to hook into the menu
Sub Rangereplication_Chat
		if iUsage<>1 then
			perform_logon
		end if
		if iUsage=1 then
			if isObject(oChatWindow) then
				oChatWindow.dispose()
			end if
			oChatWindow=CreateUnoDialog( DialogLibraries.Rangereplication.Chat )
			sChatUrl=thisComponent.URL
			oChatWindow.setVisible(true)
			oChatWindow.GetControl("TextField1").Text="Chat..."
			pull_initial_chat
		end if
end sub

' Method to hook into the X of Chat
Sub Rangereplication_Chat_close
	if isObject(oChatWindow) then
		oChatWindow.dispose()
	end if
	sChatUrl=""
end sub

' Method to hook into the menu
Sub Rangereplication_Logoff
	dim i
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if isWatchings(i)=true then
				msgbox "Can't log off while still watching"
				exit sub
			end if
		end if
	next i
	' ok, we can
	iUsage=0
	sUsername=""
	sPassword=""
end sub

' Method to hook into the menu (under more...)
Sub Rangereplication_Check
	if thisComponent.Sheets.hasByName("ggglobal_shadow") then
		' allow several attempts for login
		if iUsage<>1 then
			perform_logon	
		end if
		check_cells_for_upload
	end if
end sub

' Method to hook into the menu
Sub Rangereplication_Edit_WaitTime
	dim oTextBox1
	DialogLibraries.LoadLibrary( "Rangereplication" )
	oDialog1 = CreateUnoDialog( DialogLibraries.Rangereplication.WaitTime )
	oTextBox1 = oDialog1.GetControl("TextField1")
	if iWaitTime>0 then
		oTextBox1.text=cstr(iWaitTime)
	else
		oTextBox1.text="2500"
	end if
	oDialog1.Execute()
end sub

' Method to hook into the menu
Sub Rangereplication_Edit_Shadow 
	dim oGlobalShadow
	dim oTextBox1, oTextBox2, oTextBox3, oCheckBox1
	if not thisComponent.Sheets.hasByName("ggglobal_shadow") then
		thisComponent.Sheets.insertNewByName("ggglobal_shadow", 3)
	end if
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	DialogLibraries.LoadLibrary( "Rangereplication" )
	oDialog1 = CreateUnoDialog( DialogLibraries.Rangereplication.Edit_Settings )
	oTextBox1 = oDialog1.GetControl("TextField1")
	oTextBox1.text=oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String
	oTextBox2 = oDialog1.GetControl("TextField2")
	' make 80 the default
	if oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String="" then
		oTextBox2.text="80"
	else
		oTextBox2.text=cstr(oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).Value)
	end if
	oTextBox3 = oDialog1.GetControl("TextField3")
	oTextBox3.text=oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String
	oCheckBox1 = oDialog1.GetControl("CheckBox1")
	if cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String)=true then
		oCheckBox1.State=1
	else
		oCheckBox1.State=0
	end if
	oDialog1.Execute()
end sub

' Method to hook into the menu
Sub Rangereplication_Show_History
	' allow several attempts for login
	if iUsage<>1 then
		perform_logon	
	end if
	if iUsage=1 then
		show_cell_history
	end if
end sub

' Method to hook into the menu (under more...)
sub Rangereplication_Show_Logon
	dim sRes, iCnt
	if len(sUsername)>0 then
		sRes=sUsername & "/"
		if len(sPassword)>0 then
			for iCnt=1 to len(sPassword)
				sRes=sRes & "*"
			next iCnt
		else
			sRes=sRes & "(no PW)"
		end if
	else
		sRes="Not logged on"
	end if
	msgbox sRes
end sub

sub Rangereplication_Logout
	' only possible if watch loop is done!
	if bLoopRunning=true then
		msgbox "Still running, can't log out"
		exit sub
	end if
	' ok, we can go
	iUsage=-1
	sUsername=""
	sPassword=""
end sub


' -------- DIALOG SUPPORT --------

sub perform_logon
	DialogLibraries.LoadLibrary( "Rangereplication" )
	oDialog1 = CreateUnoDialog( DialogLibraries.Rangereplication.Logon )
	oDialog1.Execute()
end sub

sub process_logon
	iUsage=1
	dim oTextBox1, oTextBox2
	oTextBox1 = oDialog1.GetControl("TextField1")
	sUsername=oTextBox1.text
	oTextBox2 = oDialog1.GetControl("TextField2")
	sPassword=oTextBox2.text
	oDialog1.endExecute()
end sub

sub deny_logon
	iUsage=-1
	oDialog1.endExecute()
end sub

sub change_waittime
	dim oTextBox1, iNewWaitTime
	oTextBox1 = oDialog1.GetControl("TextField1")
	iNewWaitTime=cint(oTextBox1.text)
	if iNewWaitTime<1000 then
		msgbox "Minimum interval is 1 second"
		exit sub
	end if
	iWaitTime=iNewWaitTime
	oDialog1.endExecute()
end sub

sub change_shadowsheet
	dim oGlobalShadow
	dim oTextBox1, oTextBox2, oTextBox3, oCheckBox1
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	oTextBox1 = oDialog1.GetControl("TextField1")
	oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String=oTextBox1.text
	oTextBox2 = oDialog1.GetControl("TextField2")
	oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).Value=cint(oTextBox2.text)
	oTextBox3 = oDialog1.GetControl("TextField3")
	oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String=oTextBox3.text
	oCheckBox1 = oDialog1.GetControl("CheckBox1")
	if oCheckBox1.State=1 then
		oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).Value=1
	else
		oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).Value=0	
	end if
	oDialog1.endExecute()
end sub

' -------- HISTORY CODE --------

Sub show_cell_history
	dim i
	dim iTgtIndex
	' (re-use logic from cell update listener)
	dim oDialog1, oListbox
	dim sAddr
	dim oGlobalShadow
	dim oSheet, oSel
	dim sRes, sResDoc, sSingleDoc, sValue, sUser, sTime
	dim iPos, iPos2, iPosEnd, iPosUser, iPosTime
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	' find component in targets
	' give up Sheet(0) also
	iTgtIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if thisComponent.URL=oTargets(i).URL then
				iTgtIndex=i
				exit for
			end if
		end if
	next i
	if iTgtIndex<0 then
		msgbox "No active replication - no history"
		exit sub
	end if
	oSheet=thisComponent.getCurrentController().ActiveSheet
	oSel=thisComponent.getCurrentController().getSelection()
	sAddr=colToLetter(oSel.RangeAddress.StartColumn+1) & (oSel.RangeAddress.StartRow+1) & ":" & colToLetter(oSel.RangeAddress.EndColumn+1) & (oSel.RangeAddress.EndRow+1)
	' always only take one cell
	if (oSheet.getCellRangeByName(sAddr).Columns.Count * oSheet.getCellRangeByName(sAddr).Columns.Count)>1 then
		msgbox "Can not show history for several cells"
		exit sub
	end if
	' cut down to this one if we have several
	if InStr(sAddr, ":")>0 then
		sAddr=left(sAddr, InStr(sAddr, ":")-1)
	end if
	DialogLibraries.LoadLibrary( "Rangereplication" )
	oDialog1 = CreateUnoDialog( DialogLibraries.Rangereplication.History )
	oListBox = oDialog1.GetControl("ListBox1")
	do while oListbox.ItemCount>0
		oListbox.removeitems(0,1)
	loop
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/" & oSheet.getName() & "." & sAddr, "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	iPos=instr(sRes, """$history"":[")
	if iPos>0 then
		iPosEnd=Instr(iPos, sRes, "]")
		iPos=InStr(iPos+12, sRes, "{")
		do while iPos>0 and iPos<iPosEnd
			iPos2=Instr(iPos, sRes, "}")
			iPos=iPos+1
			sSingleDoc=Mid(sRes, iPos, iPos2-iPos)
			sValue=obtain_value_from_doc(sSingleDoc)
			' (for summary: abbreviate)
			if len(sValue)>45 then
				sValue=left(sValue, 27) & "..."
			end if
			iPosUser=Instr(sSingleDoc, """user"":""")
			if iPosUser>0 then
				sUser=" (" & mid(sSingleDoc, iPosUser+8, instr(iPosUser+8, sSingleDoc, """")-(iPosUser+8)) & ")"
			else
				sUser=""
			end if
			' finally also time (thx 2 https://forum.openoffice.org/en/forum/viewtopic.php?f=13&t=606)
			iPosTime=Instr(sSingleDoc, """ts"":")
			if iPosTime>0 then
			    sTime=" " & Format(cdbl(mid(sSingleDoc, iPosTime+5, instr(iPosTime+6, sSingleDoc, "}")-(iPosTime+5)))/86400000+25569, "DD.MM.YYYY HH:MM:SS")
			else
				sTime=""
			end if
			oListbox.additem(sValue & sUser & sTime, 0)
			' (and check for loop)
			iPos=Instr(iPos2+1, sRes, "{")
		loop
		oDialog1.Execute()
	else
		msgbox "No history (yet) available"
	end if
End Sub

' -------- INNER REPLICATION CODE --------

sub listen_for_changes(byref cmp)
	dim i
	dim iLstIndex
	if cmp is nothing then
		cmp=thisComponent
	end if
	iLstIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if oTargets(i).URL=cmp.URL then
				iLstIndex=i
				exit for
			end if
		end if
	next i
	if iLstIndex<0 then
		iLstIndex=iNumTargets
		oTargets(iLstIndex)=cmp
		iNumTargets=iNumTargets+1
	end if
	' jetzt der Hauptteil
	if bLoopRunning=true and bIsWatchings(iLstIndex)=true then
		msgbox "already watching"
		exit sub
	end if
	bStopWatchings(iLstIndex)=false
	bIsWatchings(iLstIndex)=true
	if isEmpty(bLoopRunning) or (not bLoopRunning=true) then
		bLoopRunning=true
		watching_loop
	end if
end sub

sub stop_listen_for_changes
	dim i
	dim iLstIndex
	iLstIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if oTargets(i).URL=thisComponent.URL then
				iLstIndex=i
				exit for
			end if
		end if
	next i
	if iLstIndex<0 then
		' no harm in just quitting
		exit sub
	end if
	bStopWatchings(iLstIndex)=true
end sub

sub pull_initial_chat
	dim oSheet, oGlobalShadow
	dim sRes, iPos, iPos2
	dim oChatContent, sChatDoc, sUser
	oChatContent=oChatWindow.GetControl("TextField1")	
	oSheet=thisComponent.getCurrentController().ActiveSheet
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")	
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_all_docs?startkey=%22Chat_%22&endkey=%22Chat_99999%22&include_docs=true", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	iPos=InStr(sRes, "{""id"":""Chat_")
	do while iPos>0
		iPos=iPos+12
		iPos2=InStr(iPos, sRes, "{""id"":""Chat_")
		if iPos2>0 then
			sChatDoc=mid(sRes, iPos, iPos2-iPos)		
		else
			sChatDoc=mid(sRes, iPos)
		end if
		rem msgbox sChatDoc
		if inStr(sChatDoc, """user"":")>0 and inStr(sChatDoc, """user"":""""")=0 then
			sUser="(by: " & decodeJson(sChatDoc, inStr(sChatDoc, """user"":")+8) & ")"
		else
			sUser="(unknown)"
		end if
		oChatContent.Text=oChatContent.Text & Chr(13) & Chr(10) & decodeJson(sChatDoc, inStr(sChatDoc, """line"":")+8) & " " & sUser		
		iPos=iPos2
	loop
end sub

sub watching_loop
	dim sSinces(100)
	dim iCnts(100) 
	dim bFirstRuns(100)
	dim iPos, iPos2
	dim sCell, bAnyChange
	dim iRevNrWe, iRevNrThey
	dim sRes, sRev, sLocalString, sRemoteString
	dim iPrevIndex, iAnchIndex
	dim oGlobalShadow
	dim oSheet, oSheetShadow
	dim oChatContent, sChatDoc, sUser
	dim sShadowInfo
	dim oInitialDialog
	' hier iterativ verwendet, nicht durch raussuchen
	dim iLstIndex
	' iLstIndex*100 ist das Offset
	dim aPreviousCells(10000), aAncientCells(10000)
	dim iPreviousCellsUsed, iAncientCellsUsed
	dim bRunAtAll, bAskRecheck
	dim cmp
	for iLstIndex=lbound(iCnts) to ubound(iCnts)
		sSinces(iLstIndex)="0"
	next iLstIndex
	for iLstIndex=lbound(iCnts) to ubound(iCnts)
		iCnts(iLstIndex)=0
	next iLstIndex
	for iLstIndex=lbound(bFirstRuns) to ubound(bFirstRuns)
		bFirstRuns(iLstIndex)=true
	next iLstIndex
	oInitialDialog=CreateUnoDialog( DialogLibraries.Rangereplication.Initcheck )
	do while true
		bRunAtAll=false
		for iLstIndex=lbound(oTargets) to ubound(oTargets)
			if isObject(oTargets(iLstIndex)) and bTargetsActive(iLstIndex)=true then
				' gracefully exit the watch loop for this index when we ran into errors
				on error goto ErrorHandler
				' check before exec to avoid errors with objects not there
				if bStopWatchings(iLstIndex)=true then
					' debug (later run forever)
					bIsWatchings(iLstIndex)=false
					rem msgbox "watch end"
				end if
				if bIsWatchings(iLstIndex)=true then
					rem if iCnt>100 or bStopWatching=true then
					if bFirstRuns(iLstIndex)=true then
						oInitialDialog.setVisible(true)
					end if
					bRunAtAll=true
					cmp=oTargets(iLstIndex)
					oGlobalShadow=cmp.Sheets.getByName("ggglobal_shadow")	
					iPreviousCellsUsed=0
					iAncientCellsUsed=0
					' all dark yellow cells should be light yellow, all light yellow white (mark diminishes after two iterations)
					' sheet is part of the name!
					for iAnchIndex=0 to 99
						if not isEmpty(aAncientCells(iLstIndex*100+iAnchIndex)) then
							if len(aAncientCells(iLstIndex*100+iAnchIndex))>0 then
								oSheet=cmp.Sheets.getByName(sheetPart(aAncientCells(iLstIndex*100+iAnchIndex)))
								oSheet.getCellRangeByName(cellPart(aAncientCells(iLstIndex*100+iAnchIndex))).getCellByPosition(0, 0).CellBackColor=RGB(255, 255, 255)
								aAncientCells(iLstIndex*100+iAnchIndex)=""
							end if
						end if
					next 
					for iPrevIndex=0 to 99
						if not isEmpty(aPreviousCells(iLstIndex*100+iPrevIndex)) then
							if len(aPreviousCells(iLstIndex*100+iPrevIndex))>0 then
								oSheet=cmp.Sheets.getByName(sheetPart(aPreviousCells(iLstIndex*100+iPrevIndex)))
								oSheet.getCellRangeByName(cellPart(aPreviousCells(iLstIndex*100+iPrevIndex))).getCellByPosition(0, 0).CellBackColor=RGB(250, 250, 160)
								aAncientCells(iLstIndex*100+iAncientCellsUsed)=aPreviousCells(iLstIndex*100+iPrevIndex)
								iAncientCellsUsed=iAncientCellsUsed+1
								aPreviousCells(iLstIndex*100+iPrevIndex)=""
							end if
						end if
					next
					' unfortunately we can't handle longpoll (no reaction to COM events / no multithreading)
					' limit 150 - as mid() is stuck with 16k+ chars
					sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_changes?since=" & sSinces(iLstIndex) & "&limit=150", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
					iPos=InStr(sRes, "{""seq"":")
					if iPos>0 then
						bAnyChange=true
					else
						bAnyChange=false
					end if
					do while iPos>0
						'msgbox iPos
						iPos=iPos+7
						iPos2=Instr(iPos, sRes, ",")
						sSinces(iLstIndex)=mid(sRes, iPos, iPos2-iPos)
						rem msgbox sSince
						iPos=InStr(iPos2+1, sRes, """id"":""")
						iPos=iPos+len("""id"":""")
						iPos2=Instr(iPos, sRes, """")
						if isChatPart(mid(sRes, iPos, iPos2-iPos)) and (not isEmpty(sChatUrl)) then
							if oTargets(iLstIndex).URL=sChatUrl then
								if isObject(oChatWindow) then
									sChatDoc=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/" & mid(sRes, iPos, iPos2-iPos), "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
									oChatContent=oChatWindow.GetControl("TextField1")	
									if inStr(sChatDoc, """user"":")>0 and inStr(sChatDoc, """user"":""""")=0 then
										sUser="(by: " & decodeJson(sChatDoc, inStr(sChatDoc, """user"":")+8) & ")"
									else
										sUser="(unknown)"
									end if
									oChatContent.Text=oChatContent.Text & Chr(13) & Chr(10) & decodeJson(sChatDoc, inStr(sChatDoc, """line"":")+8) & " " & sUser
								end if			
							end if
						elseif isSheetPart(mid(sRes, iPos, iPos2-iPos)) then
							oSheet=cmp.Sheets.getByName(sheetPart(mid(sRes, iPos, iPos2-iPos)))
							' shadow sheet, create if not exists
							if not cmp.Sheets.hasByName(oSheet.getName() & "_shadow") then
								cmp.Sheets.insertNewByName(oSheet.getName() & "_shadow", 3)
							end if
							oSheetShadow=cmp.Sheets.getByName(oSheet.getName() & "_shadow")	
							' (and cell)
							sCell=cellPart(mid(sRes, iPos, iPos2-iPos))
							rem msgbox sCell
							iPos=Instr(iPos2+1, sRes, """rev"":""")
							iPos=iPos+7
							iPos2=Instr(iPos, sRes, """")
							sRev=mid(sRes, iPos, iPos2-iPos)
							rem msgbox sCell & " - " & sRev
							if mid(sRes, iPos2, len("""}],""deleted"":true"))="""}],""deleted"":true" then
								' (ignore for now)
								rem msgbox "deleted"
							else
								' check change - ALWAYS independently of shadow sheet - check outside of RevNrWe/They (and check all combinations)
								iRevNrThey=cint(left(sRev, Instr(sRev, "-")-1)
								sShadowInfo=oSheetShadow.getCellRangeByName(sCell).getCellByPosition(0, 0).String
								if len(sShadowInfo)>0 then
									'cut off the checksum
									if InStr(sShadowInfo, "|")>0 then
										sShadowInfo=mid(sShadowInfo, InStr(sShadowInfo, "|")+1)
									end if
									iRevNrWe=cint(left(sShadowInfo, Instr(sShadowInfo, "-")-1))
								else
									' when the shadow sheet is empty, we have "0" (and we'll always be smaller than incoming)
									iRevNrWe=0
								end if
								rem msgbox iRevNrWe & " - " & iRevNrThey
								if iRevNrWe<iRevNrThey then
									rem msgbox "update"
									' when there is a local update not published yet, we can't continue
									if check_update_cell(sCell, oSheet, oSheetShadow, oGlobalShadow)=true then
										' assemble a good info for decision!
										sLocalString=oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).String
										sRemoteString=obtain_update_cell_value(sCell, sRev, oSheet, oGlobalShadow)
										' overwrite when user says so (otherwise do nothing)
										if msgbox("Conflict in " & sCell & chr(13) & "Local: " & sLocalString & chr(13) & "Remote: " & sRemoteString & chr(13) & chr(13) & "Put local value in comment and overwrite?", 4)=6 then
											local_to_comment sCell, oSheet
											if not bFirstRuns(iLstIndex) then
												aPreviousCells(iLstIndex*100+iPreviousCellsUsed)=oSheet.getName() & "." & sCell
												iPreviousCellsUsed=iPreviousCellsUsed+1
											end if
											update_cell(sCell, sRev, bFirstRuns(iLstIndex), oSheet, oSheetShadow, oGlobalShadow)
										end if
									else
										if not bFirstRuns(iLstIndex) then
											aPreviousCells(iLstIndex*100+iPreviousCellsUsed)=oSheet.getName() & "." & sCell
											iPreviousCellsUsed=iPreviousCellsUsed+1
										end if
										update_cell(sCell, sRev, bFirstRuns(iLstIndex), oSheet, oSheetShadow, oGlobalShadow)
									end if
								end if
							end if
						end if
						' finally
						iPos=InStr(iPos2+1, sRes, "{""seq"":")
					loop
					iCnts(iLstIndex)=iCnts(iLstIndex)+1
					rem if iCnt>100 or bStopWatching=true then
					if bStopWatchings(iLstIndex)=true then
						' debug (later run forever)
						bIsWatchings(iLstIndex)=false
						rem msgbox "watch end"
					end if
					' (finally): are we still first change? (incoming can be more than 150!)
					if bAnyChange=true and bFirstRuns(iLstIndex)=true then
						rem msgbox "more to get!"
						' Idee: iLstIndex 1 zurück - und unmittelbar hier nochmal
						iLstIndex=iLstIndex-1
					else
						' initial update is done - from now on we have to wait
						' now is also the chance to offer recheck (but only once)
						bAskRecheck=bFirstRuns(iLstIndex)
						bFirstRuns(iLstIndex)=false
						oInitialDialog.setVisible(false)
						' now is also the chance to offer recheck (but only once)
						if bAskRecheck=true then
							if msgbox("Did you make offline changes you now wish to share?", 4)=6 then
								Rangereplication_Check
							end if
						end if
					end if
				end if
				Goto Proceed ' rough times ask for rough measures
				ErrorHandler: 
					unassign_listeners(oTargets(iLstIndex))
					bIsWatchings(iLstIndex)=false
					Msgbox "Error during watch of '" & oTargets(iLstIndex).getURL() & "' (stopped listinging/watching): " & Error
					Resume Proceed ' reset err
				Proceed:
				rem msgbox "@: " & Err
			end if
		next iLstIndex
		if not bRunAtAll=true then
			bLoopRunning=false
			exit sub
		end if
		rem wait 10000
		if iWaitTime<=0 then
			' this timeframe makes sense as changes are visible as such
			iWaitTime=2500
		end if
		wait iWaitTime
	loop
end sub

' method for retrieving remote String (only), used above for conflicts
function obtain_update_cell_value(byval sCell, byval sRev, byref oSheet, byref oGlobalShadow)
	dim sResDoc
	sResDoc=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/cellupd/_show/cell/" & oSheet.getName() & "." & sCell, "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	obtain_update_cell_value=obtain_value_from_doc(sResDoc)
end function
' use for others also
function obtain_value_from_doc(byval sResDoc)
	dim sType
	if inStr(sResDoc, """type"":")>0 then
		sType=mid(sResDoc, inStr(sResDoc, """type"":")+8, instr(inStr(sResDoc, """type"":")+9, sResDoc, """")-inStr(sResDoc, """type"":")-8)
	else
		sType="TEXT"
	end if
	if sType="VALUE" then
		if inStr(sResDoc, """formatted"":")>0 then
			' need to handle quotations acc. to JSON (start nach ")
			obtain_value_from_doc=decodeJson(sResDoc, inStr(sResDoc, """formatted"":")+13)
		else
			obtain_value_from_doc=mid(sResDoc, inStr(sResDoc, """value"":")+8, instr(inStr(sResDoc, """value"":")+9, sResDoc, ",")-inStr(sResDoc, """value"":")-8)
		end if
	elseif sType="EMPTY" then
		obtain_value_from_doc=""
	elseif sType="FORMULA" then
		if inStr(sResDoc, """formulares"":")>0 then
			' need to handle quotations acc. to JSON (start nach ")
			obtain_value_from_doc=decodeJson(sResDoc, inStr(sResDoc, """formulares"":")+14)
		else
			obtain_value_from_doc="--unavailable--"
		end if
	else
		' need to handle quotations acc. to JSON (start nach ")
		obtain_value_from_doc=decodeJson(sResDoc, inStr(sResDoc, """value"":")+9)		
	end if
end function

' (and this one does the actual update)
sub update_cell(byval sCell, byval sRev, byval bFirstRun, byref oSheet, byref oSheetShadow, byref oGlobalShadow)
	dim sResDoc, sType, sValue, sFormulaRes, lCheck
	dim iPos, iPos2
	sResDoc=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/cellupd/_show/cell/" & oSheet.getName() & "." & sCell, "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	if inStr(sResDoc, """type"":")>0 then
		sType=mid(sResDoc, inStr(sResDoc, """type"":")+8, instr(inStr(sResDoc, """type"":")+9, sResDoc, """")-inStr(sResDoc, """type"":")-8)
	else
		sType="TEXT"
	end if
	' and update (highlighting the cell) - depending on type
	oSheetShadow.getCellRangeByName(sCell).getCellByPosition(0, 0).String="updating"
	if sType="VALUE" then
		sValue=mid(sResDoc, inStr(sResDoc, """value"":")+8, instr(inStr(sResDoc, """value"":")+9, sResDoc, ",")-inStr(sResDoc, """value"":")-8)
		oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).Value=cdbl(replace(sValue, ".", mid(cstr(1.2), 2, 1)))
		lCheck=Checksum(sValue)
	elseif sType="EMPTY" then
		oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).String=""
		lCheck=Checksum("") '=0
	elseif sType="FORMULA" then
		' need to handle quotations acc. to JSON (start nach ")
		sValue=decodeJson(sResDoc, inStr(sResDoc, """value"":")+9)
		oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).Formula=sValue
		' Checksum of Formula AND result
		sFormulaRes=""
		if inStr(sResDoc, """formulares"":")>0 then
			sFormulaRes=decodeJson(sResDoc, inStr(sResDoc, """formulares"":")+14)
		end if
		lCheck=Checksum(sValue+"###"+sFormulaRes)
	else
		' need to handle quotations acc. to JSON (start nach ")
		sValue=decodeJson(sResDoc, inStr(sResDoc, """value"":")+9)
		oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).String=sValue
		lCheck=Checksum(sValue)
	end if
	if not bFirstRun then
		oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).CellBackColor=RGB(250, 250, 0)
	end if
	oSheetShadow.getCellRangeByName(sCell).getCellByPosition(0, 0).String=lCheck & "|" & sRev
end sub

function isChatPart(byval sAddr)
	if Left(sAddr, 5)="Chat_" then
		isChatPart=true
	else
		isChatPart=false
	end if
end function

function isSheetPart(byval sAddr)
	if Instr(sAddr, ".")>0 then
		isSheetPart=true
	else
		isSheetPart=false
	end if
end function

function sheetPart(byval sAddr)
	sheetPart=left(sAddr, Instr(sAddr, ".")-1)
end function

function cellPart(byval sAddr)
	cellPart=right(sAddr, len(sAddr)-Instr(sAddr, "."))
end function

function decodeJson(byval val, byval iFromIndex)
	' vorrobben: Backslash oder Anführungszeichen ab - bis Anführungszeichen
	dim sRes
	dim iBSIndex, iQTindex, iCurrIndex, iCnt
	sRes=""
	iCurrIndex=iFromIndex
	iCnt=0
	do while true
		iQTindex=instr(iCurrIndex, val, """")
		iBSindex=instr(iCurrIndex, val, "\")
		if iQTindex=0 and iBSindex=0 then
			sRes=sRes & mid(val, iCurrIndex, len(val)-iCurrIndex+1)
			decodeJson=sRes
			exit function
		elseif iQTindex<iBSindex or (iQTindex>0 and iBSindex=0) then
			sRes=sRes & mid(val, iCurrIndex, iQTindex-iCurrIndex)
			decodeJson=sRes
			exit function
		else
			sRes=sRes & mid(val, iCurrIndex, iBSIndex-iCurrIndex)
			'if unicode, handle differently
			if mid(val, iBSIndex+1, 1)="u" then
				sRes=sRes & chr(clng("&H" & mid(val, iBSIndex+2, 4)))
				iCurrIndex=iBSIndex+6
			elseif mid(val, iBSIndex+1, 1)="r" then
				' \r = vbcr = 13
				sRes=sRes & chr(13)
				iCurrIndex=iBSIndex+2
			elseif mid(val, iBSIndex+1, 1)="n" then
				' \n = vblf = 10
				sRes=sRes & chr(10)
				iCurrIndex=iBSIndex+2
			elseif mid(val, iBSIndex+1, 1)="t" then
				' \t = 9
				sRes=sRes & chr(9)
				iCurrIndex=iBSIndex+2
			else
				' (all others)
				sRes=sRes & mid(val, iBSIndex+1, 1)
				iCurrIndex=iBSIndex+2
			end if
		end if
		' safety measure
		iCnt=iCnt+1
		if iCnt>10000 then
			decodeJson="---ERROR---"
			exit function
		end if
	loop
	'dead code
	decodeJson="---ERROR---"
end function


sub assign_listeners(byref cmp)
	dim i
	dim iTgtIndex
	dim oSheet, oDocView, oSel
	if cmp is nothing then
		cmp=thisComponent
	end if
	' keine Shadow-Sheets
	if right(cmp.getCurrentController().ActiveSheet.getName(), len("_shadow"))="_shadow" then
		msgbox "can't listen on shadow sheets"
		exit sub
	end if
	' OK, go
	' find component in targets
	' give up Sheet(0) also
	iTgtIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if cmp.URL=oTargets(i).URL then
				iTgtIndex=i
				exit for
			end if
		end if
	next i
	if iTgtIndex<0 then
		iTgtIndex=iNumTargets
		oTargets(iTgtIndex)=cmp
		bTargetsActive(iTgtIndex)=true
		iNumTargets=iNumTargets+1
	end if
	' jetzt können wir
	if isobject(oCellListeners(iTgtIndex)) or isobject(oSelListeners(iTgtIndex)) then
		msgbox "already listening"
		exit sub
	end if
	oSheet=cmp.getCurrentController().ActiveSheet
	oCellListeners(iTgtIndex)=CreateUnoListener( "CELL_", "com.sun.star.chart.XChartDataChangeEventListener" )
	oSheet.addChartDataChangeEventListener(oCellListeners(iTgtIndex))
	oSel=cmp.getCurrentController().getSelection()
	sCurrSels(iTgtIndex)=colToLetter(oSel.RangeAddress.StartColumn+1) & (oSel.RangeAddress.StartRow+1) & ":" & colToLetter(oSel.RangeAddress.EndColumn+1) & (oSel.RangeAddress.EndRow+1)
	sPrevSels(iTgtIndex)=sCurrSels(iTgtIndex)
	oSelListeners(iTgtIndex)=CreateUnoListener( "SEL_", "com.sun.star.view.XSelectionChangeListener" )
	oDocView=cmp.getCurrentController()
	oDocView.addSelectionChangeListener(oSelListeners(iTgtIndex))
	rem msgbox "listening"
end sub

sub unassign_listeners(byref cmp)
	dim i
	dim iTgtIndex
	dim oSheet, oDocView
	if cmp is nothing then
		cmp=thisComponent
	end if
	' find component in targets
	' give up Sheet(0) also
	iTgtIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if cmp.URL=oTargets(i).URL then
				iTgtIndex=i
				exit for
			end if
		end if
	next i
	if iTgtIndex<0 then
		' nothing to do
		rem msgbox "nothing to do"
		exit sub
	end if
	' jetzt können wir
	oSheet=cmp.getCurrentController().ActiveSheet
	if isObject(oCellListeners(iTgtIndex)) then 
		oSheet.removeChartDataChangeEventListener(oCellListeners(iTgtIndex))
	else
		msgbox "cell listener not removed!"
	end if
	oDocView=cmp.getCurrentController()
	if isObject(oSelListeners(iTgtIndex)) then 
		oDocView.removeSelectionChangeListener(oSelListeners(iTgtIndex))
	else
		msgbox "sel listener not removed!"
	end if
	oCellListeners(iTgtIndex)=Null
	oSelListeners(iTgtIndex)=Null
	' finally also the target
	bTargetsActive(iTgtIndex)=false
	rem msgbox "no more listening"
end sub

sub SEL_selectionChanged(oEvent)
	dim i
	dim iTgtIndex
	dim sMySel
	dim oSel
	dim oSheet
	' find component in targets
	' give up Sheet(0) also
	iTgtIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if thisComponent.URL=oTargets(i).URL then
				iTgtIndex=i
				exit for
			end if
		end if
	next i
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' keine Shadow-Sheets!
	if right(oSheet.getName(), len("_shadow"))="_shadow" then
		exit sub
	end if
	if iTgtIndex<0 then
		' silently ignore - as removeListener does not work
		'msgbox "Event but no target - impossible, ERROR!"
		exit sub
	end if
	oSel=thisComponent.getCurrentController().getSelection()
	' when choosing a filter, the selection might be gone (so check for interface)
	' msgbox oSel.dbg_SupportedInterfaces
	' msgbox oSel.dbg_Properties
	if not HasUnoInterfaces(oSel, "com.sun.star.sheet.XCellRangeAddressable") then
		' msgbox oSel.dbg_SupportedInterfaces
		' msgbox "gone"
		exit sub
	end if
	sMySel=colToLetter(oSel.RangeAddress.StartColumn+1) & (oSel.RangeAddress.StartRow+1) & ":" & colToLetter(oSel.RangeAddress.EndColumn+1) & (oSel.RangeAddress.EndRow+1)
	if sCurrSels(iTgtIndex)=sMySel then
		exit sub
	end if
	lSelTimers(iTgtIndex)=Timer()
	sPrevSels(iTgtIndex)=sCurrSels(iTgtIndex)
	sCurrSels(iTgtIndex)=sMySel
end sub

function colToLetter(byval iCol)
	dim iRemCol
	dim sRes
	iRemCol=iCol
	sRes=""
	do
		sRes = mid("ZABCDEFGHIJKLMNOPQRSTUVWXY", (iRemCol mod 26)+1, 1) & sRes
		if iRemCol>26 then
			iRemCol=fix(iRemCol/26)
		else
			iRemCol=0
		end if
	loop while iRemCol>0
	colToLetter=sRes 
end function

sub CELL_chartDataChanged(oEvent)
	dim i
	dim iTgtIndex
	dim sAddr
	dim oSheet
	' find component in targets
	' give up Sheet(0) also
	iTgtIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if thisComponent.URL=oTargets(i).URL then
				iTgtIndex=i
				exit for
			end if
		end if
	next i
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' keine Shadow-Sheets!
	if right(oSheet.getName(), len("_shadow"))="_shadow" then
		exit sub
	end if
	' (ok, jetzt zählt's)
	if iTgtIndex<0 then
		' silently ignore - as removeListener does not work
		'msgbox "Event but no target - impossible, ERROR!"
		exit sub
	end if
	' jetzt können wir
	rem msgbox colToLetter(oEvent.StartColumn+1) & (oEvent.StartRow+1) & ":" & colToLetter(oEvent.EndColumn+1) & (oEvent.EndRow+1)
	' using the event does not work - hence we need to track the sel!
	if (Timer()-lSelTimers(iTgtIndex))>=1 then
		' when 1s or more has passed since sel change, we can be sure it's actually our sel
		chart_data_upload(sCurrSels(iTgtIndex))
	else
		' when this is not the case, events might intermix, we should check both
		chart_data_upload(sPrevSels(iTgtIndex))
		chart_data_upload(sCurrSels(iTgtIndex))
	end if
end sub

sub local_to_comment(byval sAddr, byref oSheet)
	dim sComment
	sComment="Conflict resolution " & date & chr(13) & chr(10) & "Previous local values:" & chr(13) & chr(10) 
	Select Case oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).Type 
	Case com.sun.star.table.CellContentType.EMPTY
		sComment=sComment & "(empty)"
	Case com.sun.star.table.CellContentType.VALUE
		sComment=sComment & "Value: " & oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).Value & chr(13) & chr(10) & "Formatted: " & oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String
	Case com.sun.star.table.CellContentType.TEXT
		sComment=sComment & "Text: " & oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String
	Case com.sun.star.table.CellContentType.FORMULA 
		sComment=sComment & "Formula: " & oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).Formula & chr(13) & chr(10) & "Result: " & oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String
	end Select
	oSheet.getAnnotations().insertNew(oSheet.getCellRangeByName(sAddr).CellAddress, sComment)
end sub

sub chart_data_upload(byval sAddr)
	dim oSheet, oRng
	dim col, row
	rem msgbox sAddr
	oSheet=thisComponent.getCurrentController().ActiveSheet
	oRng=oSheet.getCellRangeByName(sAddr).RangeAddress
	for col=oRng.StartColumn+1 to oRng.EndColumn+1
		for row=oRng.StartRow+1 to oRng.EndRow+1
			chart_data_upload_single(colToLetter(col) & row)
		next
	next
end sub

sub chart_data_upload_single(byval sAddr)
	dim oGlobalShadow
	dim oSheet, oSheetShadow
	dim sReq, sRes, sRev, sUser, sPlatformComma, sNumFormat, sValue, sFormulaRes, lCheck
	rem msgbox sAddr
	' jetzt können wir
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' keine Shadow-Sheets!
	if right(oSheet.getName(), len("_shadow"))="_shadow" then
		exit sub
	end if
	' shadow sheet, create if not exists
	if not thisComponent.Sheets.hasByName(oSheet.getName() & "_shadow") then
		thisComponent.Sheets.insertNewByName(oSheet.getName() & "_shadow", 3)
	end if
	oSheetShadow=thisComponent.Sheets.getByName(oSheet.getName() & "_shadow")
	' nur Update, wenn nicht beide leer
	if len(oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String)=0 and len(oSheetShadow.getCellRangeByName(sAddr).getCellByPosition(0, 0).String)=0 then
		exit sub
	end if
	' und auch nur Update, wenn es nicht schon gleich ist
	if not check_update_cell(sAddr, oSheet, oSheetShadow, oGlobalShadow)=true then
		exit sub
	end if
	' OK, go
	rem msgbox sAddr
	sRev=""
	if len(oSheetShadow.getCellRangeByName(sAddr).getCellByPosition(0, 0).String)>0 then
		if oSheetShadow.getCellRangeByName(sAddr).getCellByPosition(0, 0).String="updating" then
			' when we get updated, avoid infinite circling
			exit sub
		end if
		' prevents updating newer, conflicts are NOT yet visible
		sRev=oSheetShadow.getCellRangeByName(sAddr).getCellByPosition(0, 0).String & """, "
		' cut off checksum
		if InStr(sRev, "|")>0 then
			sRev=mid(sRev, InStr(sRev, "|")+1)
		end if
		sRev="""_rev"": """ & sRev
	end if
	sUser="""user"": """ & sUsername & """"
	sPlatformComma=mid(cstr(22.33), 3, 1)
	Select Case oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).Type 
	Case com.sun.star.table.CellContentType.EMPTY
		lCheck=Checksum("") '=0
		sReq="{ " & sRev & """type"": ""EMPTY"", ""value"": null, " & sUser & " }"
	Case com.sun.star.table.CellContentType.VALUE
		sNumFormat=thisComponent.NumberFormats.getByKey(oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).NumberFormat).FormatString
		rem msgbox sNumFormat
		sValue=replace(cstr(oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).Value), sPlatformComma, ".")
		lCheck=Checksum(sValue)
		sReq="{ " & sRev & """type"": ""VALUE"", ""value"": " & sValue & ", ""formatted"": """ & prepareJson(oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String) & """, ""format"": """ & prepareJson(sNumFormat) & """, " & sUser & " }"
	Case com.sun.star.table.CellContentType.TEXT
		sValue=oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String
		lCheck=Checksum(sValue)
		sReq="{ " & sRev & """type"": ""TEXT"", ""value"": """ & prepareJson(sValue) & """, " & sUser & " }"
	Case com.sun.star.table.CellContentType.FORMULA 
		sValue=oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).Formula
		sFormulaRes=oSheet.getCellRangeByName(sAddr).getCellByPosition(0, 0).String
		lCheck=Checksum(sValue+"###"+sFormulaRes)
		sReq="{ " & sRev & """type"": ""FORMULA"", ""value"": """ & prepareJson(sValue) & """, ""formulares"": """ & prepareJson(sFormulaRes) & """, " & sUser & " }"
	end Select
	rem msgbox sReq
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/cellupd/_update/cell/" & oSheet.getName() & "." & sAddr, "PUT", "application/json", sReq, sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	rem msgbox sRes
	' check for successful update (and make conflicts visible)
	if inStr(sRes, """ok"":true")>0 then
		if inStr(sRes, """rev"":")>0 then
			sRev=mid(sRes, inStr(sRes, """rev"":")+7, instr(inStr(sRes, """rev"":")+8, sRes, """")-inStr(sRes, """rev"":")-7)
			oSheetShadow.getCellRangeByName(sAddr).getCellByPosition(0, 0).String=lCheck & "|" & sRev
		end if
	else
		msgbox "Error during update: " & sRes
	end if
end sub

sub submit_chat
	dim sChatText
	dim iLstIndex, cmp, oGlobalShadow
	dim sRes, sReq, sUser, iNextSeq, sNextSeq
	sChatText=oChatWindow.GetControl("TextField2").Text
	oChatWindow.GetControl("TextField2").Text=""
	sChatText=trim(sChatText)
	if len(sChatText)<1 then
		exit sub
	end if
	' analog Cell-Upload 
	for iLstIndex=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(iLstIndex)) and bTargetsActive(iLstIndex)=true then
			if oTargets(iLstIndex).URL=sChatUrl then
				cmp=oTargets(iLstIndex)
				exit for
			end if
		end if
	next
	if not isObject(cmp) then
		msgbox "No target for Chat submit!"
		exit sub
	end if
	oGlobalShadow=cmp.Sheets.getByName("ggglobal_shadow")
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_all_docs?descending=true&startkey=%22Chat_99999%22&endkey=%22Chat_%22&limit=1", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	if InStr(sRes, """id"":""Chat_")>0 then
		iNextSeq=cint(Mid(sRes, InStr(sRes, """id"":""Chat_")+11, 5))+1
	else
		iNextSeq=1
	end if
	sNextSeq=""+iNextSeq
	do while len(sNextSeq)<5
		sNextSeq="0"+sNextSeq
	loop
	sUser="""user"": """ & sUsername & """"
	sReq="{" & sUser & ", ""line"": """ & prepareJson(sChatText) & """}"
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/Chat_" & sNextSeq, "PUT", "application/json", sReq, sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	' check for successful update (and make conflicts visible)
	if not inStr(sRes, """ok"":true")>0 then
		msgbox "Error during chat update: " & sRes
	end if
end sub

function check_update_cell(byval sCell, byref oSheet, byref oSheetShadow, byref oGlobalShadow)
	dim sPlatformComma
	dim sRev, sValue, sFormulaRes, lCheck, lCheckComp
	' include version from shadow-sheet, if any - so we can determine whether there is a local, unreplicated update
	lCheckComp=0 'nothing or empty
	if len(oSheetShadow.getCellRangeByName(sCell).getCellByPosition(0, 0).String)>0 then
		' prevents overwrite of local updates, conflicts need to be made visible by calling fkt
		' don't talk, just KISS: diff checksum = local change!
		' no checksum (whether empty or not given) is always a change, we'll survive
		sRev=oSheetShadow.getCellRangeByName(sCell).getCellByPosition(0, 0).String
		if InStr(sRev, "|")>0 then
			lCheckComp=clng(left(sRev, InStr(sRev, "|")-1))
		else
			' when we have a rev but no checksum: sync anyway
			lCheckComp=-1
		end if
	end if	
	sPlatformComma=mid(cstr(22.33), 3, 1)
	Select Case oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).Type 
	Case com.sun.star.table.CellContentType.EMPTY
		lCheck=Checksum("") '=0
	Case com.sun.star.table.CellContentType.VALUE
		sValue=replace(cstr(oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).Value), sPlatformComma, ".")
		lCheck=Checksum(sValue)
	Case com.sun.star.table.CellContentType.TEXT
		sValue=oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).String
		lCheck=Checksum(sValue)
	Case com.sun.star.table.CellContentType.FORMULA 
		sValue=oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).Formula
		sFormulaRes=oSheet.getCellRangeByName(sCell).getCellByPosition(0, 0).String
		lCheck=Checksum(sValue+"###"+sFormulaRes)
	end Select
	if lCheck=lCheckComp then
		check_update_cell=false
	else
		check_update_cell=true
	end if
end function
 
sub check_cells_for_upload
	' walk over all cells where we have content in the sheet, but not in the shadow sheet (and publish those up)
	' next: if we have a revision, check if they match - and whether we can publish or there is a conflict to be resolved
	dim iTgtIndex
	dim i, j
	dim sCheckAddr
	dim oSheet, oSheetShadow, oGlobalShadow, oCursor, oProgressDialog, oProgressBarModel
	dim iCntRows, iCntCols
	dim bWasWatching
	' find component in targets
	' give up Sheet(0) also
	iTgtIndex=-1
	for i=lbound(oTargets) to ubound(oTargets)
		if isObject(oTargets(i)) and bTargetsActive(i)=true then
			if thisComponent.URL=oTargets(i).URL then
				iTgtIndex=i
				exit for
			end if
		end if
	next i
	if iTgtIndex<0 then
		msgbox "No target - impossible, ERROR!"
		exit sub
	end if
	' jetzt können wir
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' keine Shadow-Sheets!
	if right(oSheet.getName(), len("_shadow"))="_shadow" then
		exit sub
	end if
	bWasWatching=bIsWatchings(iTgtIndex)
	bIsWatchings(iTgtIndex)=false
	' shadow sheet, create if not exists
	if not thisComponent.Sheets.hasByName(oSheet.getName() & "_shadow") then
		thisComponent.Sheets.insertNewByName(oSheet.getName() & "_shadow", 3)
	end if
	oSheetShadow=thisComponent.Sheets.getByName(oSheet.getName() & "_shadow")
	oCursor=oSheet.createCursor()
	oCursor.GotoEndOfUsedArea(false)
	oProgressDialog=CreateUnoDialog( DialogLibraries.Rangereplication.Recheck )
	oProgressDialog.setVisible(true)
	iCntRows=oCursor.RangeAddress.EndRow+1
	iCntCols=oCursor.RangeAddress.EndColumn+1
	oProgressBarModel = oProgressDialog.getModel().getByName("ProgressBar1")
	oProgressBarModel.setPropertyValue("ProgressValueMax", iCntCols*iCntRows)
	for i=1 to iCntCols
		for j=1 to iCntRows
			oProgressBarModel.setPropertyValue("ProgressValue", (i-1)*iCntRows+(j-1))
			oProgressDialog.GetControl("Label1").Text=cstr((i-1)*iCntRows+(j-1)) & " / " & cstr(iCntCols*iCntRows)
			sCheckAddr=colToLetter(i) & j
			' (method could be simplified - but it's very readable as it is)
			if len(oSheet.getCellRangeByName(sCheckAddr).getCellByPosition(0, 0).String)>0 and len(oSheetShadow.getCellRangeByName(sCheckAddr).getCellByPosition(0, 0).String)=0 then
				' per se, there is content to INSERT
				chart_data_upload(sCheckAddr, iTgtIndex)
			elseif len(oSheet.getCellRangeByName(sCheckAddr).getCellByPosition(0, 0).String)=0 and len(oSheetShadow.getCellRangeByName(sCheckAddr).getCellByPosition(0, 0).String)>0 then
				' delete (i.e. set to empty)
				' upload function checks for diff by itself!
				chart_data_upload(sCheckAddr, iTgtIndex)
			elseif len(oSheet.getCellRangeByName(sCheckAddr).getCellByPosition(0, 0).String)>0 and len(oSheetShadow.getCellRangeByName(sCheckAddr).getCellByPosition(0, 0).String)>0 then
				' compare what is there
				' upload function checks for diff by itself!
				chart_data_upload(sCheckAddr, iTgtIndex)
			else
				' (beide leer - nix)
			end if
		next j
	next i
	oProgressDialog.dispose()
	bIsWatchings(iTgtIndex)=bWasWatching
end sub

sub show_ical_form
	DialogLibraries.LoadLibrary( "Rangereplication" )
	oDialog2 = CreateUnoDialog( DialogLibraries.Rangereplication.Create_iCal )
	oDialog2.Execute()
end sub

sub install_ical
	dim sSubjectCol, sDateCol, sStartRow
	dim oSheet, oGlobalShadow
	dim sRes, sRev, sDesign
	' pull values from form
	sSubjectCol = oDialog2.GetControl("TextField1").text
	sDateCol = oDialog2.GetControl("TextField3").text
	if len(sSubjectCol)<1 or len(sDateCol)<1 then
		msgbox "All fields must be filled"
		exit sub
	end if
	oDialog2.endExecute()
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' check if there is an iCal already on the couch
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_all_docs?startkey=%22_design/ical%22&endkey=%22_design/ical%22", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	sRev=""
	if InStr(sRes, "_design/ical")> 0 then
		sRev=mid(sRes, inStr(sRes, """rev"":")+7, instr(inStr(sRes, """rev"":")+8, sRes, """")-inStr(sRes, """rev"":")-7)
		sRev="   ""_rev"": """ & sRev & """, "
	end if
	sDesign= _ 
	"{" & _
	"   ""_id"": ""_design/ical""," & sRev & _
	"   ""views"": {" & _
	"       ""ical"": {" & _
	"           ""map"": ""function(doc){\n  if(doc._id.substring(0, \""" & oSheet.Name & "." & sSubjectCol & "\"".length)==\""" & oSheet.Name & "." & sSubjectCol & "\"" && doc.value!=null && doc.value.length>0){\n    emit(doc._id.substring(\""" & oSheet.Name & "." & sSubjectCol & "\"".length), {\""" & sSubjectCol & "\"": (doc.formatted!=null?doc.formatted:(doc.formulares!=null?doc.formulares:doc.value)), \""uid\"": doc._id});\n  }else if(doc._id.substring(0, \""" & oSheet.Name & "." & sDateCol & "\"".length)==\""" & oSheet.Name & "." & sDateCol & "\"" && doc.type==\""VALUE\"" && doc.value!=null){\n    var year=Math.floor(doc.value/365.25);\n    var month=1;\n    var months=[31, ((year%4)==0?29:28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];\n    var day=Math.floor(doc.value-year*365.25);\n    for(var i=0; i<months.length; i++){if(day>months[i]){month++;day-=months[i];}}\n    emit(doc._id.substring(\""" & oSheet.Name & "." & sDateCol & "\"".length), {\""" & sDateCol & "\"": (\""\""+(1900+year))+(\""\""+(month<9?\""0\"":\""\"")+month)+(\""\""+(day<9?\""0\"":\""\"")+day)});\n  }\n}""," & _
	"           ""reduce"": ""function(keys, values, rereduce){\n  if (rereduce){\n    return null;\n  } else {\n    if(values.length!=2){\n      return null;\n    }\n    var caption=null, date=null, uid=null;\n    for(var i=0; i<2; i++){\n      if(values[i]." & sSubjectCol & "){\n        caption=values[i]." & sSubjectCol & ";\n        uid=values[i].uid;\n      }else if(values[i]." & sDateCol & "){\n        date=values[i]." & sDateCol & ";\n      }\n    }\n    return {\""caption\"": caption, \""date\"": date, \""uid\"": uid};\n  }\n}""" & _
	"       }" & _
	"   }," & _
	"   ""lists"": {" & _
	"       ""icallist"": ""function(head, req){\n  start({\n    'headers': {\n      'Content-Type': 'text/calendar'\n    }\n  });\n  send(\""BEGIN:VCALENDAR\\nPRODID:-//Mozilla.org/NONSGML Mozilla Calendar V1.1//EN\\nVERSION:2.0\\nBEGIN:VTIMEZONE\\nTZID:Europe/Berlin\\nX-LIC-LOCATION:Europe/Berlin\\nBEGIN:DAYLIGHT\\nTZOFFSETFROM:+0100\\nTZOFFSETTO:+0200\\nTZNAME:CEST\\nDTSTART:19700329T020000\\nRRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3\\nEND:DAYLIGHT\\nBEGIN:STANDARD\\nTZOFFSETFROM:+0200\\nTZOFFSETTO:+0100\\nTZNAME:CET\\nDTSTART:19701025T030000\\nRRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10\\nEND:STANDARD\\nEND:VTIMEZONE\\n\"");\n  var row;\n  while(row=getRow()){\n    if(row.value!=null){\n      send(\""BEGIN:VEVENT\\nCREATED:\""+row.value.date+\""T000000Z\\nLAST-MODIFIED:\""+row.value.date+\""T000000Z\\nDTSTAMP:\""+row.value.date+\""T000000Z\\nUID:\""+row.value.uid+\""\\nSUMMARY:\""+row.value.caption+\""\\nDTSTART;TZID=Europe/Berlin:\""+row.value.date+\""T000000\\nDTEND;TZID=Europe/Berlin:\""+row.value.date+\""T000000\\nTRANSP:OPAQUE\\nEND:VEVENT\\n\"");\n    }\n  }\n  send(\""END:VCALENDAR\"");\n  return \""\"";\n}""" & _
	"   }" & _
	"}"
	'msgbox sDesign
	'inputbox "Design", "Design", sDesign
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/ical", "PUT", "application/json", sDesign, sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	'msgbox sRes
	if inStr(sRes, """ok"":true")>0 then
		show_ical_addr
	else
		msgbox "Error: " & sRes
	end if
end sub

sub show_ical_addr
	dim oSheet, oGlobalShadow
	dim sRes, sUrl
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' show the couch address to the iCal stream (if there is one)
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_all_docs?startkey=%22_design/ical%22&endkey=%22_design/ical%22", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	if InStr(sRes, "_design/ical")<1 then
		msgbox "No iCal stream in place!"
		exit sub
	end if
	InputBox "Address of the iCal stream: ", "iCal", "http://" & oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String & ":" & oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String & "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/ical/_list/icallist/ical?reduce=true&group=true"
end sub

function check_db_connection
	dim oGlobalShadow, sRes
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	on error goto ErrorHandler
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	on error goto 0	
	if InStr(sRes, """error"":")>0 then
		msgbox "Can't connect to DB: " & sRes
		' for unauthorized, we can retry outright
		if InStr(sRes, "unauthorized")>0 then
			Rangereplication_Logout
			Rangereplication_Start_Listening
		end if
		check_db_connection=false
		exit function
	end if
	check_db_connection=true
	exit function
	ErrorHandler:
	msgbox "Can't connect to DB: " & Error()
	on error goto 0
	check_db_connection=false
end function

sub install_cellupd
	dim oGlobalShadow
	dim sRes, sRev, sDesign, iVersion
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	' check if there is the cellupd doc already on the couch
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_all_docs?startkey=%22_design/cellupd%22&endkey=%22_design/cellupd%22&include_docs=true", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	sRev=""
	iVersion=1
	if InStr(sRes, "_design/cellupd")> 0 then
		' check whether we have at least the required version (none=1)
		if InStr(sRes, """version"":")> 0 then
			iVersion=cint(mid(sRes, inStr(sRes, """version"":")+10, instr(inStr(sRes, """version"":")+11, sRes, ",")-inStr(sRes, """version"":")-10))
		end if
		' get the rev (now if we have one!)
		sRev="""_rev"": """ & mid(sRes, Instr(sRes, """rev"":""")+7, Instr(Instr(sRes, """rev"":""")+8, sRes, """")-(Instr(sRes, """rev"":""")+7)) & ""","
		'msgbox sRev
	end if
	' we require a version, if we have it: quit
	if iVersion>=3 then
		exit sub
	end if
	' (otherwise go on)
	sDesign= _ 
	"{" & _
	"   ""_id"": ""_design/cellupd""," & sRev & _
	"   ""version"": 3," & _
	"   ""shows"": {" & _
	"       ""cell"": ""function(doc, req){\n  var ret={};\n  for(var a in doc){\n    if(a==\""_id\"" || a==\""_rev\"" || (a.substring(0, 1)!=\""_\"" && a.substring(0, 1)!=\""$\"")){\n      ret[a]=doc[a];\n    }\n  }\n  return JSON.stringify(ret);\n}""" & _
	"   }," & _
	"   ""updates"": {" & _
	"       ""cell"": ""function(doc, req){\n  var body=JSON.parse(req.body);\n  if(doc==null){\n    doc={\""_id\"": (req.id?req.id:req.uuid)};\n  }else if(doc._rev!=body._rev){\n    return [null, JSON.stringify({\""error\"":\""conflict\"",\""reason\"":\""Document update conflict.\""})];\n  }\n  var newhist={};\n  for(var a in doc){\n    if(a.substring(0, 1)!=\""_\"" && a.substring(0, 1)!=\""$\"" && (!body.hasOwnProperty(a))){\n      delete doc[a];\n    }\n  }\n  for(var a in body){\n    if(a.substring(0, 1)!=\""_\"" && a.substring(0, 1)!=\""$\""){\n      doc[a]=body[a];\n      newhist[a]=body[a];\n    }\n  }\n  newhist[\""ts\""]=(new Date()).getTime();\n  if(!doc[\""$history\""]){\n    doc[\""$history\""]=[];\n  }\n  doc[\""$history\""].push(newhist);\n  return [doc, JSON.stringify({\""ok\"":true,\""id\"":doc._id,\""rev\"":\""<X-Couch-Update-NewRev>\""})];\n}""" & _
	"   }" & _
	"}"
	'msgbox sDesign
	'inputbox "Design", "Design", sDesign
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/cellupd", "PUT", "application/json", sDesign, sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	'msgbox sRes
	if inStr(sRes, """ok"":true")>0 then
		' nothing to do
	else
		msgbox "Error: " & sRes
	end if
end sub

sub install_webview
	dim oGlobalShadow, sRes, sReplicaSource
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	' determine source: local server first, then cloudant
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/webview_install/_all_docs?startkey=%22webview%22&endkey=%22webview%22", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	if InStr(sRes, "webview")>=1 then
		' the local couch has a DB with the webview
		sReplicaSource="webview_install"
	else
		' TODO: more alternatives
		msgbox "No webview available - pls. install manually"
		exit sub
	end if
	' now replicate
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/_replicate", "POST", "application/json", "{""source"": """+sReplicaSource+""", ""target"": """+oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String+"""}", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	'msgbox sRes
	if inStr(sRes, """ok"":true")>0 then
		show_webview_addr
	else
		msgbox "Error: " & sRes
	end if
end sub

sub show_webview_addr
	dim oSheet, oGlobalShadow
	dim sRes, sUrl
	oGlobalShadow=thisComponent.Sheets.getByName("ggglobal_shadow")
	oSheet=thisComponent.getCurrentController().ActiveSheet
	' show the couch address to the iCal stream (if there is one)
	sRes=obtainViaHttp(oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String, oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String, "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_all_docs?startkey=%22webview%22&endkey=%22webview%22", "GET", "", "", sUsername, sPassword, cbool(oGlobalShadow.getCellRangeByName("A5").getCellByPosition(0, 0).String))
	if InStr(sRes, "webview")<1 then
		msgbox "No webview in place!"
		exit sub
	end if
	' TODO: include login DB in URL when available
	InputBox "Address of the webview: ", "webview", "http://" & oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String & ":" & oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String & "/" & oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String & "/_design/showfkt/_show/htmlout/webview"
end sub

sub asana_connect_EXPERIMENT
	dim sRes, sApikey, sTag, sTagId, sIdCol, sNameCol, sNotesCol, sDueCol, sParentCol
	dim iPos, iPos2, iPosPart, iRow, sId, sVal
	dim oSheet, oCursor, iCntRows, i, iFileNum
	dim oConnection
	oConnection=CreateObject("WinHttp.WinHttpRequest.5.1")
	sApikey=""
	if dir(environ("APPDATA") & "\replica\a.txt")<>""  then
		if msgbox("Use existing API-key?", 4)=6 then
			iFileNum=freefile
			open (environ("APPDATA") & "\replica\a.txt") for input as #iFileNum
			line input #iFileNum sApikey
			close #iFileNum
		end if
	end if
	if sApikey="" then
		sApiKey=InputBox("API-key?", "asana connect", "")
		if sApiKey="" then
			exit sub
		end if
		if dir(environ("APPDATA") & "\replica", 16)="" then
			mkdir environ("APPDATA") & "\replica"
		end if
		iFileNum=freefile
		open (environ("APPDATA") & "\replica\a.txt") for output as #iFileNum
		print #iFileNum sApikey
		close #iFileNum
	end if
	rem msgbox sApikey
	sTag=InputBox("Tag?", "asana connect", "")
	if sTag="" then
		exit sub
	end if
	sIdCol=InputBox("Column 4 Task ID?", "asana connect", "A")
	if sIdCol="" then
		exit sub
	end if
	sNameCol=InputBox("Column 4 Task name?", "asana connect", "B")
	if sNameCol="" then
		exit sub
	end if
	sNotesCol=InputBox("Column 4 Task description?", "asana connect", "C")
	if sNotesCol="" then
		exit sub
	end if
	sDueCol=InputBox("Column 4 due date?", "asana connect", "D")
	if sDueCol="" then
		exit sub
	end if
	sParentCol=InputBox("Column 4 parent name?", "asana connect", "E")
	if sParentCol="" then
		exit sub
	end if
	' First: resolve tag
	' actually pull result
	oConnection.Open("GET", "https://app.asana.com/api/1.0/tags?opt_fields=this.id,this.name", false)
	oConnection.SetRequestHeader("Authorization", "Basic " & base64(sApikey & ":"))
	oConnection.SetRequestHeader("Connection", "keep-alive")
	oConnection.Send()
	sRes=oConnection.ResponseText
	if left(sRes, 8)<>"{""data"":" then
		msgbox "No data returned - instead: " & sRes
		exit sub
	end if
	iPos=InStr(sRes, """name"":""" & sTag & """"
	if iPos<1 then
		Msgbox "Tag " & sTag & " not found"
		exit sub
	end if
	' now go to { to search ID from there
	while mid(sRes, iPos, 1)<>"{"
		iPos=iPos-1
	wend
	iPos=InStr(iPos, sRes, """id"":")
	iPos2=iPos+5
	while isNumeric(mid(sRes, iPos2, 1))
		iPos2=iPos2+1
	wend
	sTagId=mid(sRes, iPos+5, iPos2-iPos-5)
	rem msgbox sTagId
	' the pull tasks for tag
	' actually pull result
	oConnection.Open("GET", "https://app.asana.com/api/1.0/tasks?tag=" & sTagId & "&opt_fields=this.id,this.name,this.notes,this.due_on,this.parent,this.parent.id,this.parent.name&limit=100", false)
	oConnection.SetRequestHeader("Authorization", "Basic " & base64(sApikey & ":"))
	oConnection.SetRequestHeader("Connection", "close")
	oConnection.Send()
	sRes=oConnection.ResponseText
	if left(sRes, 8)<>"{""data"":" then
		msgbox "No data returned - instead: " & sRes
		exit sub
	end if
	oSheet=thisComponent.getCurrentController().ActiveSheet
	oCursor=oSheet.createCursor()
	oCursor.GotoEndOfUsedArea(false)
	iCntRows=oCursor.RangeAddress.EndRow+1
	iPosPart=8
	while iPosPart>0
		iPos=InStr(iPosPart, sRes, """id"":")
		iPos2=iPos+5
		while isNumeric(mid(sRes, iPos2, 1))
			iPos2=iPos2+1
		wend
		sId=mid(sRes, iPos+5, iPos2-iPos-5)
		rem msgbox sId
		' determine row from 1 to range (as in recheck), extract attrs, write them
		iRow=0
		for i=1 to iCntRows
			if oSheet.getCellRangeByName(sIdCol & i).getCellByPosition(0, 0).String=sId then
				iRow=i
				exit for
			end if
		next i
		if iRow=0 then
			do
				iRow=iRow+1
			loop while (oSheet.getCellRangeByName(sIdCol & iRow).getCellByPosition(0, 0).String<>"" or _
						oSheet.getCellRangeByName(sNameCol & iRow).getCellByPosition(0, 0).String<>"" or _
						oSheet.getCellRangeByName(sNotesCol & iRow).getCellByPosition(0, 0).String<>"" or _
						oSheet.getCellRangeByName(sDueCol & iRow).getCellByPosition(0, 0).String<>"" or _
						oSheet.getCellRangeByName(sParentCol & iRow).getCellByPosition(0, 0).String<>"")
		end if
		if iRow=1 and iCntRows<=1 then
			oSheet.getCellRangeByName(sIdCol & iRow).getCellByPosition(0, 0).String="Task ID"
			oSheet.getCellRangeByName(sNameCol & iRow).getCellByPosition(0, 0).String="Task name"
			oSheet.getCellRangeByName(sNotesCol & iRow).getCellByPosition(0, 0).String="Task description"
			oSheet.getCellRangeByName(sDueCol & iRow).getCellByPosition(0, 0).String="due date"
			oSheet.getCellRangeByName(sParentCol & iRow).getCellByPosition(0, 0).String="parent name"
			iRow=2
		end if
		oSheet.getCellRangeByName(sIdCol & iRow).getCellByPosition(0, 0).String=sId
		' (other values)
		iPos=InStr(iPosPart, sRes, """name"":""")
		sVal=decodeJson(sRes, iPos+8)
		oSheet.getCellRangeByName(sNameCol & iRow).getCellByPosition(0, 0).String=sVal
		iPos=InStr(iPosPart, sRes, """notes"":""")
		iPos=iPos+9
		if mid(sRes, iPos, 2)="""""" then 
			sVal=""
		else
			sVal=decodeJson(sRes, iPos)
		end if
		oSheet.getCellRangeByName(sNotesCol & iRow).getCellByPosition(0, 0).String=sVal
		iPos=InStr(iPosPart, sRes, """due_on"":")
		iPos=iPos+9
		if mid(sRes, iPos, 4)="null" then
			sVal=""	
		else
			sVal=decodeJson(sRes, iPos+1)
		end if
		oSheet.getCellRangeByName(sDueCol & iRow).getCellByPosition(0, 0).String=sVal
		iPos=InStr(iPosPart, sRes, """parent"":{")
		' only if the parent is with us
		iPos2=InStr(iPosPart+20, sRes, ",{")
		if iPos>0 and (iPos2<1 or iPos<iPos2) then
			iPos=iPos+9
			if mid(sRes, iPos, 4)="null" then
				sVal=""	
			else
				iPos=InStr(iPos, sRes, """name"":""")
				sVal=decodeJson(sRes, iPos+8)
			end if
		else
			sVal=""
		end if
		oSheet.getCellRangeByName(sParentCol & iRow).getCellByPosition(0, 0).String=sVal
		' next
		iPosPart=InStr(iPosPart+20, sRes, ",{")
	wend
end sub

function prepareJson(byval val)
	val=replace(val, "\", "\\")
	val=replace(val, """", "\""")
	' \r = vbcr = 13
	val=replace(val, chr(13), "\r")
	' \n = vblf = 10
	val=replace(val, chr(10), "\n")
	' \t = 9
	val=replace(val, chr(9), "\t")
	prepareJson=val
end function

function base64(byval val)
	dim offset
	dim a1, a2, a3
	dim b1, b2, b3, b4
	dim res
	offset=1
	res=""
	while offset <= len(val)
		a1=asc(mid(val, offset, 1))
		if (len(val)-offset)>0 then
			a2=asc(mid(val, offset+1, 1))
		else
			a2=0
		end if
		if (len(val)-offset)>1 then
			a3=asc(mid(val, offset+2, 1))
		else
			a3=0
		end if
		b1=int(a1/4)
		b2=(a1 mod 4)*16+int(a2/16)
		b3=(a2 mod 16)*4+int(a3/64)
		b4=a3 mod 64
		res=res & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b1+1, 1)
		res=res & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b2+1, 1)
		if (len(val)-offset)>0 then
			res=res & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b3+1, 1)
		else
			res=res & "="
		end if
		if (len(val)-offset)>1 then
			res=res & mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b4+1, 1)
		else
			res=res & "="
		end if
		offset=offset+3
	wend	
	base64=res
end function


' finally for good: call python which is bundled with libreoffice/OOo all the time and can do HTTPS (at least on mac, it does not verify certs and display a pointless warning about JRE that can be ignored)
function obtainViaHttp(byval host, byval port, byval path, byval method, byval contentType, byval content, byval user, byval password, byval ssl)
	'just call python, clng port (and package python in the lib - no matter the warning)	
	'as in https://ask.libreoffice.org/en/question/52125/execute-a-python-macro-function-from-base-and-return-value-to-base/
	dim scpr, scmod, res
	dim a(8), b(0), c(0) as variant 
	' host, port, path, method='GET', contentType=None, content=None, user=None, password=None, ssl=False
    a(0)=host
    a(1)=clng(port)
    a(2)=path
    a(3)=method
    a(4)=contentType
    a(5)=content
    a(6)=user
    a(7)=password
    a(8)=ssl
	scpr=thisComponent.getScriptProvider
	scmod = scpr.getScript("vnd.sun.star.script:SpreadsheetInSync.oxt|Scripts|httpclient.py$httpclient?language=Python&location=user:uno_packages")
	res=scmod.invoke(a, b, c)
	rem msgbox res
	obtainViaHttp=res
end function

' chunked and apache is killing us right now - so use WinHTTP as fallback until we find a real solution or go to python anyway
function obtainViaHttp_WINHTTP_OLD(byval host, byval port, byval path, byval method, byval contentType, byval content, byval user, byval password, byval ssl)
	dim oConnection, bReuseConnection, sConnectionUrl, iConnectionIndex, i
	dim sReq, sXCouchId, sXCouchUpdateNewRev, sHttpRes, sHeaders, nXPos
	dim cCR, cLF
	cCR=chr(13)
	cLF=chr(10)
	iConnectionIndex=-1
	if ssl=true then
		sConnectionUrl="https://" & host & ":" & port
	else
		sConnectionUrl="http://" & host & ":" & port
	end if
	for i=lbound(oConnections) to ubound(oConnections)
		if isObject(oConnections(i)) and sConnectionUrl=sConnectionUrls(i) and bConnectionUsed(i)<>true then
			iConnectionIndex=i
			if iConnectionUseCount(i)>50 then
				bReuseConnection=false
			else
				oConnection=oConnections(i)
				bReuseConnection=true
			end if
			exit for
		end if
	next i
	if bReuseConnection<>true or iConnectionIndex<0 then
		oConnection=CreateObject("WinHttp.WinHttpRequest.5.1")
		if iConnectionIndex<0 then
			for i=lbound(oConnections) to ubound(oConnections)
				if not isObject(oConnections(i)) then
					iConnectionIndex=i
					exit for
				end if
			next i
		end if
		if iConnectionIndex<0 then
			msgbox "Error getting a connection from the pool"
		end if
		oConnections(iConnectionIndex)=oConnection
		sConnectionUrls(iConnectionIndex)=sConnectionUrl
		iConnectionUseCount(iConnectionIndex)=0
	end if
	rem msgbox iConnectionIndex
	bConnectionUsed(iConnectionIndex)=true
	iConnectionUseCount(iConnectionIndex)=iConnectionUseCount(iConnectionIndex)+1
	oConnection.Open(method, sConnectionUrl & path, false)
	' just username is OK
	if len(user)>0 and len(password)>0 then
		oConnection.SetRequestHeader("Authorization", "Basic " & base64(user & ":" & password))
	end if
	oConnection.SetRequestHeader("Connection", "keep-alive")
	if len(contentType)>0 then
		oConnection.SetRequestHeader("Content-type", contentType)
	end if
  	if len(content)>0 then
  		oConnection.Send(content)
 	else
 		oConnection.Send()
  	end if
	sHttpRes=oConnection.ResponseText
  	' extract header values to replace in the body (be able to handle update functions)
  	sHeaders="" & oConnection.getAllResponseHeaders() & cCr & cLF
	nXPos=InStr(sHeaders, "X-Couch-Id: ")
	if nXPos>0 then
		nXPos=nXPos+len("X-Couch-Id: ")
		sXCouchId=trim(mid(sHeaders, nXPos, InStr(nXPos, sHeaders, cCR & cLF)-nXPos))
		rem msgbox "X-Couch-Id: " & sXCouchId
		sHttpRes=replace(sHttpRes, "<X-Couch-Id>", sXCouchId)
	end if
	nXPos=InStr(sHeaders, "X-Couch-Update-NewRev: ")
	if nXPos>0 then
	    nXPos=nXPos+len("X-Couch-Update-NewRev: ")
		sXCouchUpdateNewRev=trim(mid(sHeaders, nXPos, InStr(nXPos, sHeaders, cCR & cLF)-nXPos))
		rem msgbox "X-Couch-Update-NewRev: " & sXCouchUpdateNewRev
		sHttpRes=replace(sHttpRes, "<X-Couch-Update-NewRev>", sXCouchUpdateNewRev)
	end if
	rem msgbox sHttpRes
	bConnectionUsed(iConnectionIndex)=false
   	obtainViaHttp=sHttpRes
end function


'http methods from
'http://www.oooforum.org/forum/viewtopic.phtml?t=17645
'and
'http://www.oooforum.org/forum/viewtopic.phtml?p=26353
function obtainViaHttp_PLAINSOCK_OLD(byval host, byval port, byval path, byval method, byval contentType, byval content, byval user, byval password)
	dim oConnector, oConnection
	dim aByteArray, nBytesRead, nHeaderPos, nXPos
	dim cCR, cLF
	dim sReq, sXCouchId, sXCouchUpdateNewRev, sHttpRes
	dim sLog
	sLog="Log"
	cCR=chr(13)
	cLF=chr(10)
	sLog=sLog & chr(13) & "Start: " & GetSystemTicks()
 	oConnector=createUnoService("com.sun.star.connection.Connector") 
	oConnection=oConnector.connect("socket,host=" & host &",port=" & port & ",tcpNoDelay=1") 
	rem msgbox oConnection.getDescription()
	sLog=sLog & chr(13) & "Connected: " & GetSystemTicks()
	sReq=method & " " & path & " HTTP/1.1" 
	sReq=sReq & cCr & cLF & "Host: " & host & ":" & port
	if len(content)>0 then
		sReq=sReq & cCr & cLF & "Content-Length: " & StringByteLen(content)
	end if
	if len(contentType)>0 then
		sReq=sReq & cCr & cLF & "Content-type: " & contentType
	end if
	' just username is OK
	if len(user)>0 and len(password)>0 then
		sReq=sReq & cCr & cLF & "Authorization: Basic " & base64(user & ":" & password)
	end if
	sReq=sReq & cCr & cLF & "Connection: close"
	sReq=sReq & cCR & cLF & cCR & cLF
  	if len(content)>0 then
  		sReq=sReq & content
  	end if
  	rem msgbox sReq
	sLog=sLog & chr(13) & "Req string ready: " & GetSystemTicks()
  	oConnection.write(StringToByteArray(sReq))
  	oConnection.flush()
	sLog=sLog & chr(13) & "Req string sent: " & GetSystemTicks()
	aByteArray=Array() 
	sHttpRes=""
	do
	   	nBytesRead=oConnection.read(aByteArray, 200)
		sLog=sLog & chr(13) & "Chunk read: " & GetSystemTicks()
		sHttpRes=sHttpRes & left(ByteArrayToString(aByteArray), nBytesRead)
   	loop while nBytesRead>0
	sLog=sLog & chr(13) & "Res string received: " & GetSystemTicks()
	rem msgbox sHttpRes
   	nHeaderPos=InStr(sHttpRes, cCR & cLF & cCR & cLF)
   	do while asc(right(sHttpRes, 1))=10 or asc(right(sHttpRes, 1))=13
   		sHttpRes=left(sHttpRes, len(sHttpRes)-1)
   	loop
   	if nHeaderPos>0 then
   		' extract header values to replace in the body (be able to handle update functions)
		nXPos=InStr(sHttpRes, "X-Couch-Id: ")
		if nXPos>0 and nXPos<nHeaderPos then
			nXPos=nXPos+len("X-Couch-Id: ")
			sXCouchId=trim(mid(sHttpRes, nXPos, InStr(nXPos, sHttpRes, cCR & cLF)-nXPos))
			rem msgbox "X-Couch-Id: " & sXCouchId
			sHttpRes=replace(sHttpRes, "<X-Couch-Id>", sXCouchId)
		end if
		nXPos=InStr(sHttpRes, "X-Couch-Update-NewRev: ")
		if nXPos>0 and nXPos<nHeaderPos then
		    nXPos=nXPos+len("X-Couch-Update-NewRev: ")
			sXCouchUpdateNewRev=trim(mid(sHttpRes, nXPos, InStr(nXPos, sHttpRes, cCR & cLF)-nXPos))
			rem msgbox "X-Couch-Update-NewRev: " & sXCouchUpdateNewRev
			sHttpRes=replace(sHttpRes, "<X-Couch-Update-NewRev>", sXCouchUpdateNewRev)
		end if
		rem msgbox sHttpRes 
   		sHttpRes=mid(sHttpRes, nHeaderPos+4, len(sHttpRes)-(nHeaderPos+4)+1)
		rem msgbox sHttpRes
   	end if
	sLog=sLog & chr(13) & "Result extracted: " & GetSystemTicks()
	rem msgbox sLog
   	oConnection=Nothing
   	oConnector=Nothing
   	rem msgbox sHttpRes
   	obtainViaHttp=sHttpRes
end function


' see http://www.oooforum.de/viewtopic.php?t=19054 and links there

'---------------------------------------- 
'   Stuff ripped out of my library 
'---------------------------------------- 


' The next four routines were previously posted 
'   http://www.oooforum.org/forum/viewtopic.php?t=6910 


' Convert an array of bytes to a string. 
' Pass in an array of bytes. 
' Each "byte" in the array is an integer value from -128 to +127. 
' The array of bytes could have come from reading 
'  from a com.sun.star.io.XInputStream. 
' This function returns a string. 
' This function is the opposite of StringToByteArray(). 
Function ByteArrayToString(aByteArray) 
	dim cBytes, nByte, iCode, i
	cBytes="" 
	For i=LBound(aByteArray) To UBound(aByteArray)
	    iCode=0
		nByte=aByteArray(i)
		nByte=ByteToInteger(nByte)
		' added UTF-8
		if nByte>=224 then
			' 3 bytes overall
			iCode=iCode+((nByte mod 16)*4096)
			i=i+1
			nByte=aByteArray(i)
			nByte=ByteToInteger(nByte)
			iCode=iCode+((nByte mod 64)*64)
			i=i+1
			nByte=aByteArray(i)
			nByte=ByteToInteger(nByte)
			iCode=iCode+(nByte mod 64)
		elseif nByte>=192 then
			' 2 bytes overall
			iCode=iCode+((nByte mod 32)*64)
			i=i+1
			nByte=aByteArray(i)
			nByte=ByteToInteger(nByte)
			iCode=iCode+(nByte mod 64)			
		else
			' just this one byte
			iCode=iCode+nByte	
		end if
		cBytes=cBytes+Chr(iCode)
	Next i 
	ByteArrayToString() = cBytes 
End Function 

' added UTF-8
Function StringByteLen(ByVal cString As String)
	dim nNumBytes, nLen, cChar, iCode, i
	nLen=Len(cString)
	nNumBytes=0
	' added UTF-8
	for i=1 to nLen
		cChar=Mid(cString, i, 1)
		iCode=Asc(cChar)
		if iCode>=2048 then
			' 3 bytes
			nNumBytes=nNumBytes+3
		elseif iCode>=128 then
			' 2 bytes
			nNumBytes=nNumBytes+2
		else
			' just one byte
			nNumBytes=nNumBytes+1
		end if
	next i
	StringByteLen()=nNumBytes
End Function

' Convert a string into an array of bytes. 
' Pass a string value to the cString parameter. 
' The function returns an array of bytes, suitable 
'  for writing to a com.sun.star.io.XOutputStream. 
' Each "byte" in the array is an integer value from -128 to +127. 
' This function is the opposite of ByteArrayToString(). 
Function StringToByteArray(ByVal cString As String)
	dim nNumBytes, nLen, cChar, nByte, iCode, iPos, i
	nLen=Len(cString)
	' added UTF-8
	nNumBytes=StringByteLen(cString)
	Dim aBytes(nNumBytes-1) As Integer 
	iPos=0
	For i=1 To nLen
		cChar=Mid(cString, i, 1)
		iCode=Asc(cChar)
		if iCode>=2048 then
			' 3 bytes
			nByte=IntegerToByte(224+((iCode-(iCode mod 4096))/4096))
			aBytes(iPos)=nByte
			iPos=iPos+1
			iCode=iCode mod 4096
			nByte=IntegerToByte(128+((iCode-(iCode mod 64))/64))
			aBytes(iPos)=nByte
			iPos=iPos+1
			iCode=iCode mod 64
			nByte=IntegerToByte(128+iCode)
			aBytes(iPos)=nByte
			iPos=iPos+1
		elseif iCode>=128 then
			' 2 bytes
			nByte=IntegerToByte(192+((iCode-(iCode mod 64))/64))
			aBytes(iPos)=nByte
			iPos=iPos+1
			iCode=iCode mod 64
			nByte=IntegerToByte(128+iCode)
			aBytes(iPos)=nByte
			iPos=iPos+1
		else
			' just one byte
			nByte=IntegerToByte(iCode)
			aBytes(iPos)=nByte
			iPos=iPos+1
		end if
	Next i
	StringToByteArray()=aBytes()
End Function 


' Convert a byte value from the range -128 to +127 into 
'  an integer in the range 0 to 255. 
' This function is the opposite of IntegerToByte(). 
Function ByteToInteger(ByVal nByte As Integer) As Integer 
	If nByte<0 Then 
		nByte=nByte+256 
	EndIf 
	ByteToInteger()=nByte
End Function 


' This function is the opposite of ByteToInteger(). 
Function IntegerToByte(byVal nByte As Integer) As Integer 
	If nByte>127 Then 
		nByte=nByte-256
	EndIf 
	IntegerToByte()=nByte 
End Function 


' self-baked checksum
function Checksum(byVal str as String) as Long
	dim res, i
	res=0
	for i=1 to len(str)
		res=((res*31)+asc(mid(str, i, 1)))
		do while res>200000000
			res=res-200000000
		loop
	next i
	checksum=res
end function

' (END)

