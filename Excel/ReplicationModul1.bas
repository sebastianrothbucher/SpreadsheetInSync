Attribute VB_Name = "ReplicationModul1"
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
Global implObject As New ReplicationKlasse1

Public Sub register_handler(control As IRibbonControl)
    implObject.register_handler
End Sub

Public Sub unregister_handler(control As IRibbonControl)
    implObject.unregister_handler
End Sub

Public Sub recheck_sheet(control As IRibbonControl)
    implObject.recheck_sheet
End Sub

Public Sub edit_settings(control As IRibbonControl)
    implObject.edit_settings
End Sub

Public Sub logoff(control As IRibbonControl)
    implObject.logoff
End Sub
