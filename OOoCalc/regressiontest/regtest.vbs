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
' run this file for the Replica regression test! (have a couch on localhost:5984 with admin party!)
' (assisted automated test - you have to click FW but all checking is done - the pareto-solution)
' recommendation: kill soffice.bin before
set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
set WshShell=CreateObject("WScript.Shell")
Set objServiceManager=CreateObject("com.sun.star.ServiceManager")
Set objDesktop=objServiceManager.createInstance("com.sun.star.frame.Desktop")
Set objDocument=objDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, Array())
Set objDispatcher=objServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
Set objFrame=objDocument.getCurrentController().Frame

msgbox "0 - Prep: new database"
WinHttpReq.Open "DELETE", "http://localhost:5984/regressiontest", false
WinHttpReq.Send 
WinHttpReq.Open "PUT", "http://localhost:5984/regressiontest", false
WinHttpReq.Send 
if not instr(WinHttpReq.ResponseText, """ok"":true")>0 then
	msgbox WinHttpReq.ResponseText
        WScript.Quit
end if
WinHttpReq.Open "PUT", "http://localhost:5984/regressiontest/Tabelle1.A2", false
WinHttpReq.Send "{""type"": ""TEXT"", ""value"": ""regression testing"", ""user"": ""sebastian""}"
if not instr(WinHttpReq.ResponseText, """ok"":true")>0 then
	msgbox WinHttpReq.ResponseText
        WScript.Quit
end if

msgbox "1 - Setup params"
' executeDispatch blocks with dialogs - i.e. start typing in parallel and wait there
WshShell.Run """"+left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))+"regtest_1_type.vbs"+"""", 0, false
objDispatcher.executeDispatch objFrame, "vnd.sun.star.script:Rangereplication.Module1.Rangereplication_Edit_Shadow?language=Basic&location=application", "", 0, Array()
set oGlobalShadow=objDocument.Sheets.getByName("ggglobal_shadow")
if oGlobalShadow is nothing then
	msgbox "Global Shadow not set up"
	WScript.Quit
end if
if not oGlobalShadow.getCellRangeByName("A2").getCellByPosition(0, 0).String="localhost" and oGlobalShadow.getCellRangeByName("A3").getCellByPosition(0, 0).String="5984" and oGlobalShadow.getCellRangeByName("A4").getCellByPosition(0, 0).String="regressiontest" then
	msgbox "Global Shadow values not correct"
	WScript.Quit
end if

msgbox "2 - type sth before start"
objDocument.getCurrentController().ActiveSheet.getCellRangeByName("B3").getCellByPosition(0, 0).String="test up"

msgbox "3 - start replication and check"
' executeDispatch blocks forever (because of the wait loop) - so we have to start via menu (but we can wait for return)
' this is GERMAN layout ;-)
WshShell.SendKeys "%x"
WScript.Sleep 50
WshShell.SendKeys "d"
WScript.Sleep 50
WshShell.SendKeys "S"
WScript.Sleep 50
WshShell.SendKeys "S"
WScript.Sleep 50
WshShell.Run """"+left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))+"regtest_3_type.vbs"+"""", 0, true
' (wait for the data to come in)
WScript.Sleep 3000
' click away auto-recheck-shet
WshShell.SendKeys "%n"
if not objDocument.getCurrentController().ActiveSheet.getCellRangeByName("A2").getCellByPosition(0, 0).String="regression testing" then
	msgbox "Inbound not correct"
	WScript.Quit
end if
' also recheck design doc install
WinHttpReq.Open "GET", "http://localhost:5984/regressiontest/_all_docs?key=%22_design/cellupd%22", false
WinHttpReq.Send
if not instr(WinHttpReq.ResponseText, """id"":""_design/cellupd""")>0 then
	msgbox "No design doc installed - couch val: "+WinHttpReq.ResponseText
        WScript.Quit
end if

msgbox "4 - Recheck sheet"
objDispatcher.executeDispatch objFrame, "vnd.sun.star.script:Rangereplication.Module1.Rangereplication_Check?language=Basic&location=application", "", 0, Array()
WinHttpReq.Open "GET", "http://localhost:5984/regressiontest/Tabelle1.B3", false
WinHttpReq.Send
if not instr(WinHttpReq.ResponseText, """value"":""test up""")>0 then
	msgbox "Recheck not correct - couch val: "+WinHttpReq.ResponseText
        WScript.Quit
end if

msgbox "5 - type sth live"
WshShell.SendKeys "{DOWN}"
WScript.Sleep 50
WshShell.SendKeys "{DOWN}"
WScript.Sleep 50
WshShell.SendKeys "{RIGHT}"
WScript.Sleep 500
WshShell.SendKeys "again"
WScript.Sleep 50
WshShell.SendKeys " we go"
Wscript.Sleep 500
WshShell.SendKeys "{ENTER}"
WScript.Sleep 2000
WinHttpReq.Open "GET", "http://localhost:5984/regressiontest/Tabelle1.B3", false
WinHttpReq.Send
if not instr(WinHttpReq.ResponseText, """value"":""again we go""")>0 then
	msgbox "Upwards not correct - couch val: "+WinHttpReq.ResponseText
        WScript.Quit
end if
' remember _rev for next
sRes=WinHttpReq.ResponseText
iPos=InStr(sRes, """_rev"":""")+Len("""_rev"":""")
iPos2=InStr(iPos+5, sRes, """")
sRevB3=Mid(sRes, iPos, iPos2-iPos)
rem msgbox sRevB3

msgbox "6 - incoming live (with wait)"
WinHttpReq.Open "PUT", "http://localhost:5984/regressiontest/Tabelle1.B3", false
WinHttpReq.Send "{""_rev"": """+sRevB3+""", ""type"":""TEXT"", ""value"":""even more"", ""user"":""someone""}"
if not instr(WinHttpReq.ResponseText, """ok"":true")>0 then
	msgbox WinHttpReq.ResponseText
        WScript.Quit
end if
WScript.Sleep 4000
if not objDocument.getCurrentController().ActiveSheet.getCellRangeByName("B3").getCellByPosition(0, 0).String="even more" then
	msgbox "Inbound not correct"
	WScript.Quit
end if

msgbox "7 - stop replication (and check typing)"
objDispatcher.executeDispatch objFrame, "vnd.sun.star.script:Rangereplication.Module1.Rangereplication_Stop_Listening?language=Basic&location=application", "", 0, Array()
WshShell.SendKeys "^{HOME}"
WScript.Sleep 50
WshShell.SendKeys "offline"
WScript.Sleep 500
WShShell.SendKeys "{ENTER}"
WScript.Sleep 2000
WinHttpReq.Open "GET", "http://localhost:5984/regressiontest/_all_docs?key=%22Tabelle1.A1%22", false
WinHttpReq.Send
if instr(WinHttpReq.ResponseText, """_id"":""Tabelle1.A1""")>0 then
	msgbox "Did not stop replicating - couch val: "+WinHttpReq.ResponseText
        WScript.Quit
end if

msgbox "Done successfully - yeah!"
' (end)
