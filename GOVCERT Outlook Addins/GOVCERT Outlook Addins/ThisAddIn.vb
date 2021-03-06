﻿' Copyright (C) 2018, CERT Gouvernemental (GOVCERT.LU) '
' Author: Jean-Paul Weber <jean-paul.weber@govcert.etat.lu> '

'This file is part of GC-Notify.'
''
'GC-Notify is free software: you can redistribute it and/or modify'
'it under the terms of the GNU General Public License as published by'
'the Free Software Foundation, either version 3 of the License, or'
'(at your option) any later version.'
''
'GC-Notify is distributed in the hope that it will be useful,'
'but WITHOUT ANY WARRANTY; without even the implied warranty of'
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the'
'GNU General Public License for more details.'
''
'You should have received a copy of the GNU General Public License'
'along with GC-Notify.  If not, see <https://www.gnu.org/licenses/>.'
'

<System.Runtime.InteropServices.ComVisible(True)>
Public Class ThisAddIn

    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

        ' Remove Keys for office 2016 '
        RemoveLoadingTimesKeys("16.0")
        ' Remove Keys for office 2013 '
        RemoveLoadingTimesKeys("15.0")
        ' Remove Keys for office 2010 '
        RemoveLoadingTimesKeys("14.0")
        ' Programmatically determine the current Outlook item '
    End Sub

    Private Sub RemoveLoadingTimesKeys(outlookVersion As String)
        ' Remove loading times -> Workaround '
        Dim outlookBase As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Office\" + outlookVersion + "\Outlook", True)
        ' If the key does not exit then the returned value is Nothing '
        If Not outlookBase Is Nothing Then
            Dim addinsKey As Microsoft.Win32.RegistryKey = outlookBase.OpenSubKey("Addins\Govcert Outlook Addins", True)
            If Not addinsKey Is Nothing Then
                For Each valueKey As String In addinsKey.GetValueNames
                    addinsKey.DeleteValue(valueKey)
                Next
                addinsKey.Close()
            End If
            ' Note: Under AddInLoadTimes are also stored loading times however they don't provoke the disabling as they store only the 3 last loading times :/  '
            outlookBase.Close()
        End If

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New GOVCERTOutlookRibbon()
    End Function

End Class
