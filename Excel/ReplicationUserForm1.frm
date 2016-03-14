VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplicationUserForm1 
   Caption         =   "Edit Settings"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   OleObjectBlob   =   "ReplicationUserForm1.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ReplicationUserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Sub CommandButton1_Click()
    Dim globalshadow As Worksheet
    Dim i
    For i = 1 To ActiveWorkbook.Sheets.Count
        If ActiveWorkbook.Sheets(i).name = "ggglobal_shadow" Then
            Set globalshadow = ActiveWorkbook.Sheets(i)
            Exit For
        End If
    Next i
    If globalshadow Is Nothing Then
        Set globalshadow = ActiveWorkbook.Sheets.Add(, ActiveWorkbook.Sheets.Item(ActiveWorkbook.Sheets.Count), 1)
        globalshadow.name = "ggglobal_shadow"
    End If
    globalshadow.Range("A2").Value = TextBox1.Text
    globalshadow.Range("A3").Value = CLng(TextBox2.Text)
    globalshadow.Range("A4").Value = TextBox3.Text
    If CheckBox1.Value = True Then
        globalshadow.Range("A5").Value = 1
    Else
        globalshadow.Range("A5").Value = 0
    End If
    Hide
End Sub
