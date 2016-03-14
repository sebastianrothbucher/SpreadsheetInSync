VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReplicationKlasse1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
' some silly code - but we collapse asap!
Private listenbook0 As Workbook
Private listenbook1 As Workbook
Private listenbook2 As Workbook
Private listenbook3 As Workbook
Private listenbook4 As Workbook
Private listenbook5 As Workbook
Private listenbook6 As Workbook
Private listenbook7 As Workbook
Private listenbook8 As Workbook
Private listenbook9 As Workbook
Private WithEvents listentab0 As Worksheet
Attribute listentab0.VB_VarHelpID = -1
Private WithEvents listentab1 As Worksheet
Attribute listentab1.VB_VarHelpID = -1
Private WithEvents listentab2 As Worksheet
Attribute listentab2.VB_VarHelpID = -1
Private WithEvents listentab3 As Worksheet
Attribute listentab3.VB_VarHelpID = -1
Private WithEvents listentab4 As Worksheet
Attribute listentab4.VB_VarHelpID = -1
Private WithEvents listentab5 As Worksheet
Attribute listentab5.VB_VarHelpID = -1
Private WithEvents listentab6 As Worksheet
Attribute listentab6.VB_VarHelpID = -1
Private WithEvents listentab7 As Worksheet
Attribute listentab7.VB_VarHelpID = -1
Private WithEvents listentab8 As Worksheet
Attribute listentab8.VB_VarHelpID = -1
Private WithEvents listentab9 As Worksheet
Attribute listentab9.VB_VarHelpID = -1
Private WithEvents listenconn0 As WinHttpRequest
Attribute listenconn0.VB_VarHelpID = -1
Private WithEvents listenconn1 As WinHttpRequest
Attribute listenconn1.VB_VarHelpID = -1
Private WithEvents listenconn2 As WinHttpRequest
Attribute listenconn2.VB_VarHelpID = -1
Private WithEvents listenconn3 As WinHttpRequest
Attribute listenconn3.VB_VarHelpID = -1
Private WithEvents listenconn4 As WinHttpRequest
Attribute listenconn4.VB_VarHelpID = -1
Private WithEvents listenconn5 As WinHttpRequest
Attribute listenconn5.VB_VarHelpID = -1
Private WithEvents listenconn6 As WinHttpRequest
Attribute listenconn6.VB_VarHelpID = -1
Private WithEvents listenconn7 As WinHttpRequest
Attribute listenconn7.VB_VarHelpID = -1
Private WithEvents listenconn8 As WinHttpRequest
Attribute listenconn8.VB_VarHelpID = -1
Private WithEvents listenconn9 As WinHttpRequest
Attribute listenconn9.VB_VarHelpID = -1
' remember sinces by index
Dim sSinces(0 To 9) As String
' Connection re-use
Dim oConnections(100)
Dim sConnectionUrls(100)
Dim bConnectionUsed(100)
Dim iConnectionUseCount(100)
' remember logon
Public bLoggedOn As Boolean
Public sUser As String
Public sPassword As String

Public Sub register_handler()
    Dim globalshadow As Worksheet
    Dim currentsheet As Worksheet
    If LCase(Right(ActiveSheet.name, Len("_shadow"))) = "_shadow" Then
        MsgBox "Can not listen on shadow sheets"
        Exit Sub
    End If
    ' some silly code - but we collapse asap! (1/2)
    If ActiveSheet Is listentab0 Or ActiveSheet Is listentab1 Or ActiveSheet Is listentab2 Or ActiveSheet Is listentab3 Or ActiveSheet Is listentab4 Or ActiveSheet Is listentab5 Or ActiveSheet Is listentab6 Or ActiveSheet Is listentab7 Or ActiveSheet Is listentab8 Or ActiveSheet Is listentab9 Then
        ' MsgBox "already listening"
        Exit Sub
    End If
    ' global shadow (poss. Setup needed)
    Set currentsheet = ActiveSheet
    If findSheet(ActiveWorkbook, "ggglobal_shadow") Is Nothing Then
        If MsgBox("No database specified yet. Specify now?", vbYesNo) = vbYes Then
            edit_settings
            currentsheet.Activate
            If MsgBox("Retry now?", vbYesNo) = vbYes Then
                register_handler
            End If
        End If
        Exit Sub
    End If
    ' have to be logged in
    If bLoggedOn = False Then
        logon
        If bLoggedOn = False Then
            Exit Sub
        End If
    End If
    ' test DB connection once!
    If check_db_connection() = False Then
        Exit Sub
    End If
    ' get design doc in place
    install_cellupd
    ' some silly code - but we collapse asap! (2/2)
    If listentab0 Is Nothing Then
        Set listenbook0 = ActiveWorkbook
        Set listentab0 = ActiveSheet
        sSinces(0) = "0"
        listen0
        'MsgBox "Listening 0"
    ElseIf listentab1 Is Nothing Then
        Set listenbook1 = ActiveWorkbook
        Set listentab1 = ActiveSheet
        sSinces(1) = "0"
        listen1
        'MsgBox "Listening 1"
    ElseIf listentab2 Is Nothing Then
        Set listenbook2 = ActiveWorkbook
        Set listentab2 = ActiveSheet
        sSinces(2) = "0"
        listen2
        'MsgBox "Listening 2"
    ElseIf listentab3 Is Nothing Then
        Set listenbook3 = ActiveWorkbook
        Set listentab3 = ActiveSheet
        sSinces(3) = "0"
        listen3
        'MsgBox "Listening 3"
    ElseIf listentab4 Is Nothing Then
        Set listenbook4 = ActiveWorkbook
        Set listentab4 = ActiveSheet
        sSinces(4) = "0"
        listen4
        'MsgBox "Listening 4"
    ElseIf listentab5 Is Nothing Then
        Set listenbook5 = ActiveWorkbook
        Set listentab5 = ActiveSheet
        sSinces(5) = "0"
        listen5
        'MsgBox "Listening 5"
    ElseIf listentab6 Is Nothing Then
        Set listenbook6 = ActiveWorkbook
        Set listentab6 = ActiveSheet
        sSinces(6) = "0"
        listen6
        'MsgBox "Listening 6"
    ElseIf listentab7 Is Nothing Then
        Set listenbook7 = ActiveWorkbook
        Set listentab7 = ActiveSheet
        sSinces(7) = "0"
        listen7
        'MsgBox "Listening 7"
    ElseIf listentab8 Is Nothing Then
        Set listenbook8 = ActiveWorkbook
        Set listentab8 = ActiveSheet
        sSinces(8) = "0"
        listen8
        'MsgBox "Listening 8"
    ElseIf listentab9 Is Nothing Then
        Set listenbook9 = ActiveWorkbook
        Set listentab9 = ActiveSheet
        sSinces(9) = "0"
        listen9
        'MsgBox "Listening 9"
    Else
        MsgBox "Can't listen to more than 10"
        Exit Sub
    End If
End Sub

Public Sub unregister_handler()
    ' some silly code - but we collapse asap!
    If ActiveSheet Is listentab0 Then
        Set listenbook0 = Nothing
        Set listentab0 = Nothing
        listenconn0.Abort
        Set listenconn0 = Nothing
    ElseIf ActiveSheet Is listentab1 Then
        Set listenbook1 = Nothing
        Set listentab1 = Nothing
        listenconn1.Abort
        Set listenconn1 = Nothing
    ElseIf ActiveSheet Is listentab2 Then
        Set listenbook2 = Nothing
        Set listentab2 = Nothing
        listenconn2.Abort
        Set listenconn2 = Nothing
    ElseIf ActiveSheet Is listentab3 Then
        Set listenbook3 = Nothing
        Set listentab3 = Nothing
        listenconn3.Abort
        Set listenconn3 = Nothing
    ElseIf ActiveSheet Is listentab4 Then
        Set listenbook4 = Nothing
        Set listentab4 = Nothing
        listenconn4.Abort
        Set listenconn4 = Nothing
    ElseIf ActiveSheet Is listentab5 Then
        Set listenbook5 = Nothing
        Set listentab5 = Nothing
        listenconn5.Abort
        Set listenconn5 = Nothing
    ElseIf ActiveSheet Is listentab6 Then
        Set listenbook6 = Nothing
        Set listentab6 = Nothing
        listenconn6.Abort
        Set listenconn6 = Nothing
    ElseIf ActiveSheet Is listentab7 Then
        Set listenbook7 = Nothing
        Set listentab7 = Nothing
        listenconn7.Abort
        Set listenconn7 = Nothing
    ElseIf ActiveSheet Is listentab8 Then
        Set listenbook8 = Nothing
        Set listentab8 = Nothing
        listenconn8.Abort
        Set listenconn8 = Nothing
    ElseIf ActiveSheet Is listentab9 Then
        Set listenbook9 = Nothing
        Set listentab9 = Nothing
        listenconn9.Abort
        Set listenconn9 = Nothing
    Else
        'MsgBox "not listening (anyway)"
    End If
End Sub

Public Sub logon()
    ReplicationLogonForm.Show vbModal
    'MsgBox sUser & ":" & sPassword
End Sub

Public Sub logoff()
    If Not (listentab0 Is Nothing And listentab1 Is Nothing And listentab2 Is Nothing And listentab3 Is Nothing And listentab4 Is Nothing And listentab5 Is Nothing And listentab6 Is Nothing And listentab7 Is Nothing And listentab8 Is Nothing And listentab9 Is Nothing) Then
        MsgBox "Can't log off: a replication is still running"
        Exit Sub
    End If
    bLoggedOn = False
    sUser = ""
    sPassword = ""
End Sub

Public Sub recheck_sheet()
    ' walk over the full sheet and compare with Checksum (poss. in up itself)
    ReplicationRecheckForm.Show vbModeless
    chg ActiveWorkbook, ActiveSheet, ActiveSheet.UsedRange
    ReplicationRecheckForm.Hide
End Sub

Private Sub Class_Initialize()
    bLoggedOn = False
    sUser = ""
    sPassword = ""
End Sub

Private Sub listentab0_Change(ByVal target As Range)
    chg listenbook0, listentab0, target
End Sub

Private Sub listentab1_Change(ByVal target As Range)
    chg listenbook1, listentab1, target
End Sub

Private Sub listentab2_Change(ByVal target As Range)
    chg listenbook2, listentab2, target
End Sub

Private Sub listentab3_Change(ByVal target As Range)
    chg listenbook3, listentab3, target
End Sub

Private Sub listentab4_Change(ByVal target As Range)
    chg listenbook4, listentab4, target
End Sub

Private Sub listentab5_Change(ByVal target As Range)
    chg listenbook5, listentab5, target
End Sub

Private Sub listentab6_Change(ByVal target As Range)
    chg listenbook6, listentab6, target
End Sub

Private Sub listentab7_Change(ByVal target As Range)
    chg listenbook7, listentab7, target
End Sub

Private Sub listentab8_Change(ByVal target As Range)
    chg listenbook8, listentab8, target
End Sub

Private Sub listentab9_Change(ByVal target As Range)
    chg listenbook9, listentab9, target
End Sub

Private Sub listen0()
    If listenconn0 Is Nothing Then
        Set listenconn0 = New WinHttpRequest
    End If
    listenIgnite 0, listenconn0, listenbook0, listentab0
End Sub

Private Sub listenconn0_OnResponseFinished()
    listenReceive 0, listenconn0, listenbook0, listentab0
    listen0
End Sub

Private Sub listenconn0_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 0: " & ErrorDescription
End Sub

Private Sub listen1()
    If listenconn1 Is Nothing Then
        Set listenconn1 = New WinHttpRequest
    End If
    listenIgnite 1, listenconn1, listenbook1, listentab1
End Sub

Private Sub listenconn1_OnResponseFinished()
    listenReceive 1, listenconn1, listenbook1, listentab1
    listen1
End Sub

Private Sub listenconn1_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 1: " & ErrorDescription
End Sub

Private Sub listen2()
    If listenconn2 Is Nothing Then
        Set listenconn2 = New WinHttpRequest
    End If
    listenIgnite 2, listenconn2, listenbook2, listentab2
End Sub

Private Sub listenconn2_OnResponseFinished()
    listenReceive 2, listenconn2, listenbook2, listentab2
    listen2
End Sub

Private Sub listenconn2_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 2: " & ErrorDescription
End Sub

Private Sub listen3()
    If listenconn3 Is Nothing Then
        Set listenconn3 = New WinHttpRequest
    End If
    listenIgnite 3, listenconn3, listenbook3, listentab3
End Sub

Private Sub listenconn3_OnResponseFinished()
    listenReceive 3, listenconn3, listenbook3, listentab3
    listen3
End Sub

Private Sub listenconn3_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 3: " & ErrorDescription
End Sub

Private Sub listen4()
    If listenconn4 Is Nothing Then
        Set listenconn4 = New WinHttpRequest
    End If
    listenIgnite 4, listenconn4, listenbook4, listentab4
End Sub

Private Sub listenconn4_OnResponseFinished()
    listenReceive 4, listenconn4, listenbook4, listentab4
    listen4
End Sub

Private Sub listenconn4_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 4: " & ErrorDescription
End Sub

Private Sub listen5()
    If listenconn5 Is Nothing Then
        Set listenconn5 = New WinHttpRequest
    End If
    listenIgnite 5, listenconn5, listenbook5, listentab5
End Sub

Private Sub listenconn5_OnResponseFinished()
    listenReceive 5, listenconn5, listenbook5, listentab5
    listen5
End Sub

Private Sub listenconn5_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 5: " & ErrorDescription
End Sub

Private Sub listen6()
    If listenconn6 Is Nothing Then
        Set listenconn6 = New WinHttpRequest
    End If
    listenIgnite 6, listenconn6, listenbook6, listentab6
End Sub

Private Sub listenconn6_OnResponseFinished()
    listenReceive 6, listenconn6, listenbook6, listentab6
    listen6
End Sub

Private Sub listenconn6_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 6: " & ErrorDescription
End Sub

Private Sub listen7()
    If listenconn7 Is Nothing Then
        Set listenconn7 = New WinHttpRequest
    End If
    listenIgnite 7, listenconn7, listenbook7, listentab7
End Sub

Private Sub listenconn7_OnResponseFinished()
    listenReceive 7, listenconn7, listenbook7, listentab7
    listen7
End Sub

Private Sub listenconn7_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 7: " & ErrorDescription
End Sub

Private Sub listen8()
    If listenconn8 Is Nothing Then
        Set listenconn8 = New WinHttpRequest
    End If
    listenIgnite 8, listenconn8, listenbook8, listentab8
End Sub

Private Sub listenconn8_OnResponseFinished()
    listenReceive 8, listenconn8, listenbook8, listentab8
    listen8
End Sub

Private Sub listenconn8_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 8: " & ErrorDescription
End Sub

Private Sub listen9()
    If listenconn9 Is Nothing Then
        Set listenconn9 = New WinHttpRequest
    End If
    listenIgnite 9, listenconn9, listenbook9, listentab9
End Sub

Private Sub listenconn9_OnResponseFinished()
    listenReceive 9, listenconn9, listenbook9, listentab9
    listen9
End Sub

Private Sub listenconn9_OnError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
    MsgBox "Listen error 9: " & ErrorDescription
End Sub

Private Sub listenIgnite(ByVal index As Integer, ByRef listenconn As WinHttpRequest, ByRef book As Workbook, ByRef sheet As Worksheet)
    Dim globalshadow As Worksheet
    If sheet Is Nothing Then
        ' over and out ;-)
        Exit Sub
    End If
    Set globalshadow = findSheet(book, "ggglobal_shadow")
    ' how "resync" for fist time
    If sSinces(index) = "0" And ReplicationInitcheckForm.Visible = False Then
        ReplicationInitcheckForm.Show vbModeless
    ElseIf ReplicationInitcheckForm.Visible = True Then
        ReplicationInitcheckForm.Hide
    End If
    listenViaHttp index, listenconn, globalshadow, sUser, sPassword
End Sub

Private Sub listenViaHttp(ByVal index As Integer, ByRef listenconn As WinHttpRequest, ByRef globalshadow As Worksheet, ByVal user, ByVal password)
    ' last param is important! (timeout>couch w/ 60000)
    listenconn.SetTimeouts 10000, 60000, 30000, 70000
    Dim prot As String
    If CBool(globalshadow.Range("A5").Value) = True Then
        prot = "https"
    Else
        prot = "http"
    End If
    listenconn.Open "GET", prot & "://" & globalshadow.Range("A2").Text & ":" & globalshadow.Range("A3").Text & "/" & globalshadow.Range("A4").Text & "/_changes?feed=longpoll&timeout=60000&since=" & sSinces(index), True
    ' just username is OK
    If Len(user) > 0 And Len(password) > 0 Then
        ' listenconn.SetCredentials user, password, 0 ' 0 = HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
        listenconn.SetRequestHeader "Authorization", "Basic " & base64(user & ":" & password)
    End If
    listenconn.Send
End Sub

Private Sub listenReceive(ByVal index As Integer, ByRef listenconn As WinHttpRequest, ByRef book As Workbook, ByRef sheet As Worksheet)
    Dim sRes, iPos, iPos2, bAnyChange, sCell, sRev, iRevNrThey, sShadowInfo, iRevNrWe, sLocalString, sRemoteString, bAskRecheck
    Dim shadowsheet As Worksheet
    If sheet Is Nothing Then
        ' over and out ;-)
        Exit Sub
    End If
    If sSinces(index) = "0" Then
        bAskRecheck = True
    Else
        bAskRecheck = False
    End If
    sRes = listenconn.ResponseText
    'MsgBox sRes
    iPos = InStr(sRes, "{""seq"":")
    If iPos > 0 Then
        bAnyChange = True
    Else
        bAnyChange = False
    End If
    Do While iPos > 0
        iPos = iPos + 7
        iPos2 = InStr(iPos, sRes, ",")
        sSinces(index) = Mid(sRes, iPos, iPos2 - iPos)
        iPos = InStr(iPos2 + 1, sRes, """id"":""")
        iPos = iPos + Len("""id"":""")
        iPos2 = InStr(iPos, sRes, """")
        If isSheetPart(Mid(sRes, iPos, iPos2 - iPos), sheet) Then
            Set shadowsheet = findSheet(book, sheet.name + "_shadow")
            If shadowsheet Is Nothing Then
                Set shadowsheet = book.Sheets.Add(, book.Sheets.Item(book.Sheets.Count), 1)
                shadowsheet.name = sheet.name + "_shadow"
                sheet.Activate
            End If
            ' (and cell)
            sCell = Mid(sRes, iPos, iPos2 - iPos)
            sCell = Mid(sCell, InStr(sCell, ".") + 1)
            iPos = InStr(iPos2 + 1, sRes, """rev"":""")
            iPos = iPos + 7
            iPos2 = InStr(iPos, sRes, """")
            sRev = Mid(sRes, iPos, iPos2 - iPos)
            'MsgBox sCell & " - " & sRev
            If Mid(sRes, iPos2, Len("""}],""deleted"":true")) = """}],""deleted"":true" Then
                ' (ignore for now)
                'MsgBox "deleted"
            Else
                ' check change - ALWAYS independently of shadow sheet
                iRevNrThey = CInt(Left(sRev, InStr(sRev, "-") - 1))
                sShadowInfo = shadowsheet.Range(sCell).Text
                ' avoid running in circles
                If Not sShadowInfo = "updating" Then
                    If Len(sShadowInfo) > 0 Then
                        'cut off the checksum
                        If InStr(sShadowInfo, "|") > 0 Then
                            sShadowInfo = Mid(sShadowInfo, InStr(sShadowInfo, "|") + 1)
                        End If
                        iRevNrWe = CInt(Left(sShadowInfo, InStr(sShadowInfo, "-") - 1))
                    Else
                        ' when the shadow sheet is empty, we have "0" (and we'll always be smaller than incoming)
                        iRevNrWe = 0
                    End If
                    'MsgBox iRevNrWe & " - " & iRevNrThey
                    If iRevNrWe < iRevNrThey Then
                        'MsgBox "update " & sCell & " - " & sRev & " - " & iRevNrWe & " - " & iRevNrThey
                        If check_update_cell(sheet, shadowsheet, sheet.Range(sCell)) = True Then
                            ' assemble a good info for decision!
                            sLocalString = sheet.Range(sCell).Text
                            sRemoteString = obtain_update_cell_value(book, sheet, sCell, sRev)
                            ' overwrite when user says so (otherwise do nothing)
                            If MsgBox("Conflict in " & sCell & vbCr & "Local: " & sLocalString & vbCr & "Remote: " & sRemoteString & vbCr & vbCr & "Put local value in comment and overwrite?", vbYesNo) = vbYes Then
                                sheet.Range(sCell).AddComment "Conflict resolution " & CStr(Now) & vbCrLf & "Previous local values:" & vbCrLf & sLocalString
                                ' TODO: coloring
                                update_cell book, sheet, shadowsheet, sCell, sRev
                            End If
                        Else
                            ' TODO: coloring
                            update_cell book, sheet, shadowsheet, sCell, sRev
                        End If ' (check_update_cell)
                    End If ' (iRevNrWe < iRevNrThey)
                End If ' ("updating")
            End If ' (not deleted)
        End If ' (isSheetPart)
        ' finally
        iPos = InStr(iPos2 + 1, sRes, "{""seq"":")
    Loop
    If bAskRecheck = True Then
        If MsgBox("Did you make offline changes you now wish to share?", vbYesNo) = vbYes Then
            recheck_sheet
        End If
    End If
End Sub

Private Function isSheetPart(ByVal sAddr, ByRef sheet As Worksheet)
    If InStr(sAddr, ".") > 0 Then
        ' only watch our OWN sheet
        If Left(sAddr, Len(sheet.name) + 1) = (sheet.name & ".") Then
            isSheetPart = True
        Else
            isSheetPart = False
        End If
    Else
        isSheetPart = False
    End If
End Function

' method for retrieving remote String (only), used above for conflicts
Private Function obtain_update_cell_value(ByRef book As Workbook, ByRef sheet As Worksheet, ByVal sCell, ByVal sRev)
    Dim globalshadow As Worksheet
    Dim sResDoc
    Set globalshadow = findSheet(book, "ggglobal_shadow")
    sResDoc = obtainViaHttp(globalshadow.Range("A2").Text, globalshadow.Range("A3").Text, "/" & globalshadow.Range("A4").Text & "/_design/cellupd/_show/cell/" & sheet.name & "." & sCell, "GET", "", "", sUser, sPassword, CBool(globalshadow.Range("A5").Value))
    obtain_update_cell_value = obtain_value_from_doc(sResDoc)
End Function
' use for others also
Private Function obtain_value_from_doc(ByVal sResDoc)
    Dim sType
    If InStr(sResDoc, """type"":") > 0 Then
        sType = Mid(sResDoc, InStr(sResDoc, """type"":") + 8, InStr(InStr(sResDoc, """type"":") + 9, sResDoc, """") - InStr(sResDoc, """type"":") - 8)
    Else
        sType = "TEXT"
    End If
    If sType = "VALUE" Then
        If InStr(sResDoc, """formatted"":") > 0 Then
            ' need to handle quotations acc. to JSON (start nach ")
            obtain_value_from_doc = decodeJson(sResDoc, InStr(sResDoc, """formatted"":") + 13)
        Else
            obtain_value_from_doc = Mid(sResDoc, InStr(sResDoc, """value"":") + 8, InStr(InStr(sResDoc, """value"":") + 9, sResDoc, ",") - InStr(sResDoc, """value"":") - 8)
        End If
    ElseIf sType = "EMPTY" Then
        obtain_value_from_doc = ""
    ElseIf sType = "FORMULA" Then
        If InStr(sResDoc, """formulares"":") > 0 Then
            ' need to handle quotations acc. to JSON (start nach ")
            obtain_value_from_doc = decodeJson(sResDoc, InStr(sResDoc, """formulares"":") + 14)
        Else
            obtain_value_from_doc = "--unavailable--"
        End If
    Else
        ' need to handle quotations acc. to JSON (start nach ")
        obtain_value_from_doc = decodeJson(sResDoc, InStr(sResDoc, """value"":") + 9)
    End If
End Function

' (and this one does the actual update)
Private Sub update_cell(ByRef book As Workbook, ByRef sheet As Worksheet, ByRef shadow As Worksheet, ByVal sCell, ByVal sRev)
    Dim globalshadow As Worksheet
    Dim sResDoc, sType, sValue, sFormulaRes, lCheck
    Dim iPos, iPos2
    Set globalshadow = findSheet(book, "ggglobal_shadow")
    sResDoc = obtainViaHttp(globalshadow.Range("A2").Text, globalshadow.Range("A3").Text, "/" & globalshadow.Range("A4").Text & "/_design/cellupd/_show/cell/" & sheet.name & "." & sCell, "GET", "", "", sUser, sPassword, CBool(globalshadow.Range("A5").Value))
    If InStr(sResDoc, """type"":") > 0 Then
        sType = Mid(sResDoc, InStr(sResDoc, """type"":") + 8, InStr(InStr(sResDoc, """type"":") + 9, sResDoc, """") - InStr(sResDoc, """type"":") - 8)
    Else
        sType = "TEXT"
    End If
    ' and update (highlighting the cell) - depending on type
    shadow.Range(sCell).Value = "updating"
    If sType = "VALUE" Then
        sValue = Mid(sResDoc, InStr(sResDoc, """value"":") + 8, InStr(InStr(sResDoc, """value"":") + 9, sResDoc, ",") - InStr(sResDoc, """value"":") - 8)
        ' be sure with platform comma here also
        sheet.Range(sCell).Value = CDbl(Replace(sValue, ".", Mid(CStr(1.2), 2, 1)))
        lCheck = checksum(sValue)
    ElseIf sType = "EMPTY" Then
        sheet.Range(sCell).Value = ""
        lCheck = checksum("") '=0
    ElseIf sType = "FORMULA" Then
        ' need to handle quotations acc. to JSON (start nach ")
        sValue = decodeJson(sResDoc, InStr(sResDoc, """value"":") + 9)
        sheet.Range(sCell).Formula = sValue
        ' Checksum of Formula AND result
        sFormulaRes = ""
        If InStr(sResDoc, """formulares"":") > 0 Then
            sFormulaRes = decodeJson(sResDoc, InStr(sResDoc, """formulares"":") + 14)
        End If
        lCheck = checksum(sValue + "###" + sFormulaRes)
    Else
        ' need to handle quotations acc. to JSON (start nach ")
        sValue = decodeJson(sResDoc, InStr(sResDoc, """value"":") + 9)
        sheet.Range(sCell).Value = sValue
        lCheck = checksum(sValue)
    End If
    ' TODO: coloring
    shadow.Range(sCell).Value = lCheck & "|" & sRev
End Sub

Private Sub chg(ByRef book As Workbook, ByRef sheet As Worksheet, ByVal target As Range)
    Dim i, j As Long
    'MsgBox "Chg " & target.Address
    ' iterate over cells and upload
    For i = 1 To target.Columns.Count
        For j = 1 To target.Rows.Count
            chg_update (i - 1) * target.Rows.Count + (j - 1), target.Columns.Count * target.Rows.Count
            up book, sheet, target.Cells(j, i)
        Next j
    Next i
End Sub
' update status win if we have one
Private Sub chg_update(ByVal pos As Long, ByVal total As Long)
    Dim i As Long
    Dim sBars As String
    If ReplicationRecheckForm.Visible = True Then
        ' check how many bars fit
        sBars = ""
        For i = 1 To Round(pos * 60 / total)
            sBars = sBars & "|"
        Next i
        ReplicationRecheckForm.Label1.Caption = sBars
        ReplicationRecheckForm.Label2.Caption = pos & " / " & total
        ReplicationRecheckForm.Repaint
    End If
End Sub

' Methode für Upload (wenn nötig!)
Private Sub up(ByRef book As Workbook, ByRef sheet As Worksheet, ByRef cell As Range)
    Dim json, resp, revin, rev As String
    Dim sPlatformComma
    Dim shadow As Worksheet
    Dim globalshadow As Worksheet
    Set shadow = findSheet(book, sheet.name + "_shadow")
    If shadow Is Nothing Then
        Set shadow = book.Sheets.Add(, book.Sheets.Item(book.Sheets.Count), 1)
        shadow.name = sheet.name + "_shadow"
        sheet.Activate
    End If
    If shadow.Range(cell.Address).Text = "updating" Then
        ' avoid running in circles
        Exit Sub
    End If
    Set globalshadow = findSheet(book, "ggglobal_shadow")
    sPlatformComma = Mid(CStr(1.5), 2, 1)
    ' nfa for old=new (includes emptyness)
    If check_update_cell(sheet, shadow, cell) = False Then
        Exit Sub
    End If
    json = "{"
    If Len(shadow.Range(cell.Address).Text) > 0 Then
        revin = shadow.Range(cell.Address).Text
        If InStr(revin, "|") > 0 Then
            revin = Mid(revin, InStr(revin, "|") + 1)
        End If
        json = json & """_rev"": """ & revin & """, "
    End If
    If Len(cell.Text) = 0 Then
        json = json & """type"": ""EMPTY"", ""value"": null"
    ElseIf Len(cell.Formula) > 0 And Left(cell.Formula, 1) = "=" Then
        json = json & """type"": ""FORMULA"", " & """value"": """ & prepareJson(cell.Formula) & """, ""formulares"": """ & prepareJson(cell.Text) & """"
    ElseIf VarType(cell.Value2) = vbBoolean Then
        json = json & """type"": ""VALUE"", ""value"": "
        If cell.Value2 = True Then
            json = json & "1"
        Else
            json = json & "0"
        End If
        json = json & ", ""formatted"": """ & prepareJson(cell.Text) & """"
    ElseIf IsNumeric(cell.Value2) Then
        ' platform comma
        json = json & """type"": ""VALUE"", ""value"": " & Replace(CStr(cell.Value2), sPlatformComma, ".") & ", ""formatted"": """ & prepareJson(cell.Text) & """"
    Else
        json = json & """type"": ""TEXT"", ""value"": """ & prepareJson(cell.Text) & """"
    End If
    json = json & ", ""user"": """ & sUser & """" & "}"
    ' avoid running in circles
    shadow.Range(cell.Address).Value = "updating"
    'MsgBox "pushing " & json
    ' (sync mode) - use design doc!
    resp = obtainViaHttp(globalshadow.Range("A2").Text, globalshadow.Range("A3").Text, "/" & globalshadow.Range("A4").Text & "/_design/cellupd/_update/cell/" & sheet.name & "." & Replace(cell.Address, "$", ""), "PUT", "application/json", json, sUser, sPassword, CBool(globalshadow.Range("A5").Value))
    If InStr(resp, """ok"":true") > 0 Then
        rev = Mid(resp, InStr(resp, """rev"":") + 7, InStr(InStr(resp, """rev"":") + 8, resp, """") - InStr(resp, """rev"":") - 7)
        shadow.Range(cell.Address).Value = checksumForCell(cell) & "|" & rev
    Else
        MsgBox "Error in: " & vbCr & json & vbCr & "Error:" & vbCr & resp
    End If
    sheet.Activate
End Sub

Public Sub edit_settings()
    Dim globalshadow As Worksheet
    Set globalshadow = findSheet(ActiveWorkbook, "ggglobal_shadow")
    If Not globalshadow Is Nothing Then
        ReplicationUserForm1.TextBox1.Text = globalshadow.Range("A2").Text
        If Len(globalshadow.Range("A3").Text) > 0 Then
            ReplicationUserForm1.TextBox2.Text = globalshadow.Range("A3").Text
        Else
            ReplicationUserForm1.TextBox2.Text = "80"
        End If
        ReplicationUserForm1.TextBox3.Text = globalshadow.Range("A4").Text
        ReplicationUserForm1.CheckBox1.Value = CBool(globalshadow.Range("A5").Value)
    Else
        ReplicationUserForm1.TextBox1.Text = ""
        ReplicationUserForm1.TextBox2.Text = "80"
        ReplicationUserForm1.TextBox3.Text = ""
    End If
    ReplicationUserForm1.Show vbModal
End Sub

Private Function base64(ByVal val As String) As String
    Dim offset
    Dim a1, a2, a3
    Dim b1, b2, b3, b4
    Dim res
    offset = 1
    res = ""
    While offset <= Len(val)
        a1 = Asc(Mid(val, offset, 1))
        If (Len(val) - offset) > 0 Then
            a2 = Asc(Mid(val, offset + 1, 1))
        Else
            a2 = 0
        End If
        If (Len(val) - offset) > 1 Then
            a3 = Asc(Mid(val, offset + 2, 1))
        Else
            a3 = 0
        End If
        b1 = Int(a1 / 4)
        b2 = (a1 Mod 4) * 16 + Int(a2 / 16)
        b3 = (a2 Mod 16) * 4 + Int(a3 / 64)
        b4 = a3 Mod 64
        res = res & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b1 + 1, 1)
        res = res & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b2 + 1, 1)
        If (Len(val) - offset) > 0 Then
            res = res & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b3 + 1, 1)
        Else
            res = res & "="
        End If
        If (Len(val) - offset) > 1 Then
            res = res & Mid("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", b4 + 1, 1)
        Else
            res = res & "="
        End If
        offset = offset + 3
    Wend
    base64 = res
End Function

Private Function obtainViaHttp(ByVal host, ByVal port, ByVal path, ByVal method, ByVal ContentType, ByVal content, ByVal user, ByVal password, ByVal useSsl As Boolean)
    Dim oConnection As WinHttpRequest
    Dim bReuseConnection, sConnectionUrl, iConnectionIndex, i
    Dim sReq, sXCouchId, sXCouchUpdateNewRev, sHttpRes, sHeaders, nXPos
    iConnectionIndex = -1
    If useSsl = True Then
        sConnectionUrl = "https://" & host & ":" & port
    Else
        sConnectionUrl = "http://" & host & ":" & port
    End If
    For i = LBound(oConnections) To UBound(oConnections)
        If IsObject(oConnections(i)) And sConnectionUrl = sConnectionUrls(i) And bConnectionUsed(i) <> True Then
            iConnectionIndex = i
            If iConnectionUseCount(i) > 50 Then
                bReuseConnection = False
            Else
                Set oConnection = oConnections(i)
                bReuseConnection = True
            End If
            Exit For
        End If
    Next i
    If bReuseConnection <> True Or iConnectionIndex < 0 Then
        Set oConnection = New WinHttpRequest
        If iConnectionIndex < 0 Then
            For i = LBound(oConnections) To UBound(oConnections)
                If Not IsObject(oConnections(i)) Then
                    iConnectionIndex = i
                    Exit For
                End If
            Next i
        End If
        If iConnectionIndex < 0 Then
            MsgBox "Error getting a connection from the pool"
        End If
        Set oConnections(iConnectionIndex) = oConnection
        sConnectionUrls(iConnectionIndex) = sConnectionUrl
        iConnectionUseCount(iConnectionIndex) = 0
    End If
    'msgbox iConnectionIndex
    bConnectionUsed(iConnectionIndex) = True
    iConnectionUseCount(iConnectionIndex) = iConnectionUseCount(iConnectionIndex) + 1
    oConnection.Open method, sConnectionUrl & path, False
    ' just username is OK
    If Len(user) > 0 And Len(password) > 0 Then
        ' oConnection.SetCredentials user, password, 0 ' 0 = HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
        oConnection.SetRequestHeader "Authorization", "Basic " & base64(user & ":" & password)
    End If
    oConnection.SetRequestHeader "Connection", "keep-alive"
    If Len(content) > 0 Then
        oConnection.Send content
    Else
        oConnection.Send
    End If
    sHttpRes = oConnection.ResponseText
    ' extract header values to replace in the body (be able to handle update functions)
    sHeaders = "" & oConnection.GetAllResponseHeaders() & vbCr & vbLf
    nXPos = InStr(sHeaders, "X-Couch-Id: ")
    If nXPos > 0 Then
        nXPos = nXPos + Len("X-Couch-Id: ")
        sXCouchId = Trim(Mid(sHeaders, nXPos, InStr(nXPos, sHeaders, vbCr & vbLf) - nXPos))
        'msgbox "X-Couch-Id: " & sXCouchId
        sHttpRes = Replace(sHttpRes, "<X-Couch-Id>", sXCouchId)
    End If
    nXPos = InStr(sHeaders, "X-Couch-Update-NewRev: ")
    If nXPos > 0 Then
        nXPos = nXPos + Len("X-Couch-Update-NewRev: ")
        sXCouchUpdateNewRev = Trim(Mid(sHeaders, nXPos, InStr(nXPos, sHeaders, vbCr & vbLf) - nXPos))
        'msgbox "X-Couch-Update-NewRev: " & sXCouchUpdateNewRev
        sHttpRes = Replace(sHttpRes, "<X-Couch-Update-NewRev>", sXCouchUpdateNewRev)
    End If
    'msgbox sHttpRes
    bConnectionUsed(iConnectionIndex) = False
    obtainViaHttp = sHttpRes
End Function

Private Function findSheet(ByRef book As Workbook, ByVal name As String)
    Dim i
    For i = 1 To book.Worksheets.Count
        If book.Worksheets.Item(i).name = name Then
            Set findSheet = book.Worksheets.Item(i)
            Exit Function
        End If
    Next i
    Set findSheet = Nothing
End Function

Private Function check_db_connection()
    Dim globalshadow As Worksheet
    Dim resp
    Set globalshadow = findSheet(ActiveWorkbook, "ggglobal_shadow")
    On Error GoTo ErrorHandler
    resp = obtainViaHttp(globalshadow.Range("A2").Text, globalshadow.Range("A3").Text, "/" & globalshadow.Range("A4").Text & "/", "GET", "", "", sUser, sPassword, CBool(globalshadow.Range("A5").Value))
    On Error GoTo 0
    If InStr(resp, """error"":") > 0 Then
        MsgBox "Can't connect to DB: " & resp
        If InStr(resp, "unauthorized") > 0 Then
            logoff
            register_handler
        End If
        check_db_connection = False
        Exit Function
    End If
    check_db_connection = True
    Exit Function
ErrorHandler:
    MsgBox "Can't connect to DB: " & Err.Source & " - " & Err.Description
    On Error GoTo 0
    check_db_connection = False
End Function

Private Sub install_cellupd()
    Dim globalshadow As Worksheet
    Dim sRes, sRev, sDesign, iVersion
    Set globalshadow = findSheet(ActiveWorkbook, "ggglobal_shadow")
    ' check if there is the cellupd doc already on the couch
    sRes = obtainViaHttp(globalshadow.Range("A2").Text, globalshadow.Range("A3").Text, "/" & globalshadow.Range("A4").Text & "/_all_docs?startkey=%22_design/cellupd%22&endkey=%22_design/cellupd%22&include_docs=true", "GET", "", "", sUser, sPassword, CBool(globalshadow.Range("A5").Value))
    sRev = ""
    iVersion = 1
    If InStr(sRes, "_design/cellupd") > 0 Then
        ' check whether we have at least the required version (none=1)
        If InStr(sRes, """version"":") > 0 Then
            iVersion = CInt(Mid(sRes, InStr(sRes, """version"":") + 10, InStr(InStr(sRes, """version"":") + 11, sRes, ",") - InStr(sRes, """version"":") - 10))
        End If
        ' get the rev (now if we have one!)
        sRev = """_rev"": """ & Mid(sRes, InStr(sRes, """rev"":""") + 7, InStr(InStr(sRes, """rev"":""") + 8, sRes, """") - (InStr(sRes, """rev"":""") + 7)) & ""","
        'msgbox sRev
    End If
    ' we require a version, if we have it: quit
    If iVersion >= 3 Then
        Exit Sub
    End If
    ' (otherwise go on)
    sDesign = _
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
    sRes = obtainViaHttp(globalshadow.Range("A2").Text, globalshadow.Range("A3").Text, "/" & globalshadow.Range("A4").Text & "/_design/cellupd", "PUT", "application/json", sDesign, sUser, sPassword, CBool(globalshadow.Range("A5").Value))
    'msgbox sRes
    If InStr(sRes, """ok"":true") > 0 Then
        ' nothing to do
    Else
        MsgBox "Error: " & sRes
    End If
End Sub

Private Function prepareJson(ByVal val As String) As String
    ' JSON library for VB? (/ Script Control?)
    val = Replace(val, "\", "\\")
    val = Replace(val, """", "\""")
    prepareJson = val
End Function

Private Function decodeJson(ByVal val, ByVal iFromIndex)
    ' vorrobben: Backslash oder Anführungszeichen ab - bis Anführungszeichen
    Dim sRes
    Dim iBSIndex, iQTindex, iCurrIndex, iCnt
    sRes = ""
    iCurrIndex = iFromIndex
    iCnt = 0
    Do While True
        iQTindex = InStr(iCurrIndex, val, """")
        iBSIndex = InStr(iCurrIndex, val, "\")
        If iQTindex = 0 And iBSIndex = 0 Then
            sRes = sRes & Mid(val, iCurrIndex, Len(val) - iCurrIndex + 1)
            decodeJson = sRes
            Exit Function
        ElseIf iQTindex < iBSIndex Or (iQTindex > 0 And iBSIndex = 0) Then
            sRes = sRes & Mid(val, iCurrIndex, iQTindex - iCurrIndex)
            decodeJson = sRes
            Exit Function
        Else
            sRes = sRes & Mid(val, iCurrIndex, iBSIndex - iCurrIndex)
            'if unicode, handle differently
            If Mid(val, iBSIndex + 1, 1) = "u" Then
                sRes = sRes & Chr(CLng("&H" & Mid(val, iBSIndex + 2, 4)))
                iCurrIndex = iBSIndex + 6
            Else
                sRes = sRes & Mid(val, iBSIndex + 1, 1)
                iCurrIndex = iBSIndex + 2
            End If
        End If
        ' safety measure
        iCnt = iCnt + 1
        If iCnt > 10000 Then
            decodeJson = "---ERROR---"
            Exit Function
        End If
    Loop
    'dead code
    decodeJson = "---ERROR---"
End Function

Private Function check_update_cell(ByRef sheet As Worksheet, ByRef shadow As Worksheet, ByRef cell As Range)
    Dim checksum
    checksum = 0
    If Len(shadow.Range(cell.Address).Text) > 0 And InStr(shadow.Range(cell.Address).Text, "|") > 0 Then
        checksum = CLng(Left(shadow.Range(cell.Address).Text, InStr(shadow.Range(cell.Address).Text, "|") - 1))
    End If
    ' nfa for old=new (includes emptyness)
    If checksumForCell(cell) = checksum Then
        check_update_cell = False
    Else
        check_update_cell = True
    End If
End Function

Private Function checksumForCell(ByRef cell As Range)
    Dim sPlatformComma
    sPlatformComma = Mid(CStr(1.5), 2, 1)
    If Len(cell.Text) = 0 Then
        checksumForCell = checksum("") '=0
    ElseIf Len(cell.Formula) > 0 And Left(cell.Formula, 1) = "=" Then
        checksumForCell = checksum(cell.Formula & "###" & cell.Text)
    ElseIf VarType(cell.Value2) = vbBoolean Then
        If cell.Value2 = True Then
            checksumForCell = checksum("1")
        Else
            checksumForCell = checksum("0")
        End If
    ElseIf IsNumeric(cell.Value2) Then
        ' platform comma
        checksumForCell = checksum(Replace(CStr(cell.Value2), sPlatformComma, "."))
    Else
        checksumForCell = checksum(cell.Text)
    End If
End Function

' self-baked checksum (same as in OOo!)
Private Function checksum(ByVal str As String) As Long
    Dim res, i
    res = 0
    For i = 1 To Len(str)
        res = ((res * 31) + Asc(Mid(str, i, 1)))
        Do While res > 200000000
            res = res - 200000000
        Loop
    Next i
    checksum = res
End Function

