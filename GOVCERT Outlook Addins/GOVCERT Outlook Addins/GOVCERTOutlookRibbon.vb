' Copyright (C) 2018, CERT Gouvernemental (GOVCERT.LU) '
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

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Net.NetworkInformation
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop

<System.Runtime.InteropServices.ComVisible(True)>
Public Class GOVCERTOutlookRibbon
    Implements IRibbonExtensibility

    Private ribbon As IRibbonUI
    Private ipAddress As String = Nothing
    Private sysInformation As String = Nothing

    Const PS_PUBLIC_STRINGS As String = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}"
    Const PR_TRANSPORT_MESSAGE_HEADERS As String = "http://schemas.microsoft.com/mapi/proptag/0x007D001E"
    Const EMAIL_HEADER_TAG_REGEX As String = "^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)"
    Public Sub New()
        ' Constructor '

    End Sub


#Region "Mail Gathering"
    Private Function GetCurrentMail() As Outlook.MailItem
        ' Returns the current mail '
        Dim inspector As Outlook.Inspector = Globals.ThisAddIn.Application.ActiveInspector()
        If TypeOf inspector Is Outlook.Inspector Then
            If TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                Dim mail As Microsoft.Office.Interop.Outlook.MailItem = CType(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
                Return mail
            End If
        End If
        Return Nothing
    End Function

    Private Function GetCurrentSelection() As Outlook.Selection
        ' Returns the current selection of mails '
        Dim selectedMails As Microsoft.Office.Interop.Outlook.Selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection
        Return selectedMails

    End Function
#End Region


#Region "Mail Processing"

    Private Function ExtractHeaderInformation(ByVal mailItem As Outlook.MailItem) As String()
        ' Returns and array of header informations '
        Dim results As New List(Of String)

        Dim headerStr As String = CType(mailItem.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS), String)
        Dim tags() As String = My.Settings.INTERESTING_HEADER_FIELDS.Split(","c)
        Dim received As New List(Of String)
        Dim temp As String
        For Each m As Match In Regex.Matches(headerStr, EMAIL_HEADER_TAG_REGEX, RegexOptions.Multiline)
            For Each needle As String In tags
                If m.Value.StartsWith(needle + ":") Then
                    temp = m.Value.Replace(System.Environment.NewLine, "").Replace(vbTab, " ").Replace(vbCrLf, "")
                    If needle = "Received" Then
                        received.Add(temp)
                    Else
                        results.Add(temp)
                    End If
                    Exit For
                End If
            Next
        Next
        If received.Count > 0 Then
            If received.Count >= 4 Then
                Dim i As Integer = 4
                Do While i > 0
                    results.Add(received.Item(received.Count - i))
                    i = i - 1
                Loop
            Else
                received.Reverse()
                For Each item In received
                    results.Add(item)
                Next
            End If
        End If

        If results.Count > 0 Then
            Return results.ToArray()
        Else
            Return Nothing
        End If

    End Function

#End Region

#Region "Information Gathering"
    Private Function GetSystemInformation() As String
        If sysInformation Is Nothing Then
            ' Returns a line of informations about the sending machine '
            Me.sysInformation = "Computername:" + System.Environment.MachineName + " (" + System.Environment.OSVersion.ToString() + ")"
        End If
        Return Me.sysInformation

    End Function

    Private Function GetNetworkIP() As String
        ' Returns the IP and network interface Informations '
        If Me.ipAddress Is Nothing Then
            Dim result As String = ""
            For Each ni As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                If ni.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Or ni.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then
                    For Each ip As UnicastIPAddressInformation In ni.GetIPProperties.UnicastAddresses
                        If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                            result += ni.Name + "(" + ip.Address.ToString() + ")" + System.Environment.NewLine
                        End If
                    Next

                End If
            Next
            If result.Length > 0 Then
                Me.ipAddress = result
            Else
                Me.ipAddress = "Could Not be determined"
            End If
        End If
        Return Me.ipAddress

    End Function

    Private Function PrepareAttachment(ByRef outgoingMail As Outlook.MailItem, ByVal attachmentMail As Outlook.MailItem) As MailDetails
        ' Attaches the mail as attachment to the mail to be sent and returns the information gathered from the mail to attach '
        Dim attachMail As Boolean = True
        If Not (attachmentMail.Subject Is Nothing) Then
            If attachmentMail.Subject.Contains(My.Settings.SPAM_TAG) Then
                ' This means that the mail was alredy tagged as spam then the user must confirm to send it '
                'This means that the mail was already tagged as spam then show the dialog for confirmation'
                Dim message As String = My.Resources.SPAMDialogText
                TemplateFiller(message, "Email", attachmentMail.SenderEmailAddress)
                TemplateFiller(message, "Subject", attachmentMail.Subject)

                Dim result As DialogResult = MessageBox.Show(message, "Email already Tagged", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    attachMail = False
                Else
                    ' Set forced header '
                    ' Set X-Headers '
                    outgoingMail.PropertyAccessor.SetProperty(PS_PUBLIC_STRINGS + "/X-GC-Notify-Force-Send", "true")
                End If
            End If
        End If
        If attachMail Then
            outgoingMail.Attachments.Add(attachmentMail, Outlook.OlAttachmentType.olByValue)
            Dim result = New MailDetails()
            result.Subject = attachmentMail.Subject
            result.From = attachmentMail.SenderEmailAddress
            result.NumberOfAttachments = attachmentMail.Attachments.Count
            result.HeaderInformations = ExtractHeaderInformation(attachmentMail)
            Return result
        Else
            Return Nothing
        End If
    End Function

    Private Sub PrepareErrorMail(e As Exception)
        Dim email As Outlook.MailItem
        email = CType(Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        email.Subject = My.Settings.SOC_NEW_MAIL_Subject + " Error Occured"
        email.To = My.Settings.SUPPORT_MAIL
        email.Body = My.Resources.ErrorMail
        TemplateFiller(email.Body, "Message", e.Message)
        TemplateFiller(email.Body, "Stacktrace", e.StackTrace)
        TemplateFiller(email.Body, "Version", My.Resources.BTN_VERSIION)
        email.Display()

    End Sub

#End Region

#Region "Mail Creation Helpers"
    Private Function IsEmptyMail(email As Outlook.MailItem) As Boolean
        ' Checks if the mail is empty '
        Dim counter As Integer = 0
        If String.IsNullOrEmpty(email.Subject) Then
            counter = counter + 1
        End If
        If String.IsNullOrEmpty(email.Body) Or email.Body = " " Then
            counter = counter + 1
        End If
        If String.IsNullOrEmpty(email.To) Then
            counter = counter + 1
        End If
        If String.IsNullOrEmpty(email.CC) Then
            counter = counter + 1
        End If
        If String.IsNullOrEmpty(email.BCC) Then
            counter = counter + 1
        End If
        If email.Attachments.Count = 0 Then
            counter = counter + 1
        End If
        Return counter = 6
    End Function

    Private Function CreateNewMail() As Outlook.MailItem
        ' Creates the new mail to send out '
        Dim email As Outlook.MailItem
        email = CType(Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem), Outlook.MailItem)
        PopulateNewMail(email)
        Return email
    End Function

    Private Sub PopulateNewMail(ByRef email As Outlook.MailItem)
        ' Sets the values for the new mail '
        email.Subject = My.Settings.SOC_MAIL_SUBJECT_TAG
        email.Body = My.Resources.NewMailBody
        email.To = My.Settings.SOC_MAIL
        If Not (String.IsNullOrEmpty(My.Settings.SOC_MAIL_CC)) Then
            email.CC = My.Settings.SOC_MAIL_CC
        End If
        If Not (String.IsNullOrEmpty(My.Settings.SOC_MAIL_BCC)) Then
            email.BCC = My.Settings.SOC_MAIL_BCC
        End If
        email.PropertyAccessor.SetProperty(PS_PUBLIC_STRINGS + "/X-GC-Notify-Version", My.Resources.BTN_VERSIION)
    End Sub

    Private Sub TemplateFiller(ByRef template As String, ByVal placeholder As String, ByVal value As String)
        ' Replaces the place holder in the template with the value converted to the correct type '
        Dim stringValue As String = value
        If stringValue Is Nothing Then
            stringValue = "Nothing could be found"
        End If
        template = template.Replace("{{" + placeholder + "}}", stringValue)
    End Sub

    Private Sub TemplateFiller(ByRef template As String, ByVal placeholder As String, ByVal value As String())
        ' Replaces the place holder in the template with the value converted to the correct type '
        Dim stringValue As String = Nothing

        If value Is Nothing Then
            stringValue = "Nothing could be found"
        Else
            stringValue = ""
            For Each item In value
                stringValue += item + System.Environment.NewLine
            Next
        End If
        template = template.Replace("{{" + placeholder + "}}", stringValue)

    End Sub

    Private Sub TemplateFiller(ByRef template As String, ByVal placeholder As String, ByVal value As Integer)
        ' Replaces the place holder in the template with the value converted to the correct type '
        Dim stringValue As String = value.ToString()
        template = template.Replace("{{" + placeholder + "}}", stringValue)
    End Sub

    Private Function CreateMailDetails(mailDetails As MailDetails, Optional counter As Integer = 1) As String
        ' Populates the Mail Details template with the details of the mailDetails '
        Dim detailsTemplate = My.Resources.EmailDetails
        TemplateFiller(detailsTemplate, "EmailCounter", counter)
        TemplateFiller(detailsTemplate, "From", mailDetails.From)
        TemplateFiller(detailsTemplate, "Subject", mailDetails.Subject)
        TemplateFiller(detailsTemplate, "HeaderDetails", mailDetails.HeaderInformations)
        TemplateFiller(detailsTemplate, "AttachmentCount", mailDetails.NumberOfAttachments)
        Return detailsTemplate
    End Function
#End Region

#Region "Mail Sending"
    Private Sub ProcessMainWindowMail()
        Try
            Dim SelectedMails As Outlook.Selection = GetCurrentSelection()
            Dim outGoingMail As Outlook.MailItem = CreateNewMail()
            Dim attachments As New List(Of MailDetails)
            Dim temp As MailDetails = Nothing
            For Each email As Outlook.MailItem In SelectedMails
                temp = PrepareAttachment(outGoingMail, email)
                If Not (temp Is Nothing) Then
                    attachments.Add(temp)
                End If
            Next
            If attachments.Count = 0 Then
                MessageBox.Show(My.Resources.NoSelectionError, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                ' Set the correct subject '
                Dim subject As String
                If attachments.Count = 1 Then
                    subject = My.Settings.SOC_MAIL_SUBJECT_TAG + " - " + attachments.Item(0).GetSubjectLine()
                Else
                    subject = My.Settings.SOC_MAIL_SUBJECT_TAG + " - Multiple Emails"
                End If
                Dim mailDetails As String = ""
                Dim counter As Integer = 0
                For Each attachment In attachments
                    counter += 1
                    mailDetails += CreateMailDetails(attachment, counter)
                Next
                ' Populate Mail '
                outGoingMail.Subject = subject
                Dim body = My.Resources.SuspectBody
                TemplateFiller(body, "attachments", mailDetails)
                TemplateFiller(body, "HostDetails", GetSystemInformation())
                TemplateFiller(body, "NetworkDetails", GetNetworkIP())
                outGoingMail.Body = body
                outGoingMail.Display()
            End If
        Catch ex As Exception
            PrepareErrorMail(ex)
        End Try

    End Sub

    Private Sub ProcessWindowMail()
        Try
            Dim email As Outlook.MailItem = GetCurrentMail()
            Dim sendMail As Boolean = True
            ' First check if it Is not a new mail '
            ' If the EntryID Is empty the mail was not send nor saved, hence this is a new one '
            If String.IsNullOrEmpty(email.EntryID) Then
                ' Check however if the email was not already populated '
                If IsEmptyMail(email) Then
                    sendMail = False
                    PopulateNewMail(email)
                    TemplateFiller(email.Body, "HostDetails", GetSystemInformation())
                    TemplateFiller(email.Body, "NetworkDetails", GetNetworkIP())

                Else
                    ' Check if it is not already a new generated mail '
                    Try
                        email.PropertyAccessor.GetProperty(PS_PUBLIC_STRINGS + "/X-GC-Notify-Version")
                        MessageBox.Show(My.Resources.ResendError, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        sendMail = False
                    Catch ex As COMException
                        ' This is then something different '
                        Dim result As DialogResult = MessageBox.Show(My.Resources.OverWriteConfirm, "Do you want to proceed?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                        If result = DialogResult.Yes Then
                            sendMail = False
                            PopulateNewMail(email)
                            TemplateFiller(email.Body, "HostDetails", GetSystemInformation())
                            TemplateFiller(email.Body, "NetworkDetails", GetNetworkIP())
                        End If
                    End Try
                End If
            End If
            If sendMail Then
                ' Populate the mail and send it also close the opened mail '
                Dim outGoingMail As Outlook.MailItem = CreateNewMail()
                Dim temp As MailDetails = PrepareAttachment(outGoingMail, email)
                Dim subject As String = My.Settings.SOC_MAIL_SUBJECT_TAG + " - " + temp.GetSubjectLine()
                Dim mailDetails As String = CreateMailDetails(temp)
                outGoingMail.Subject = subject
                Dim body = My.Resources.SuspectBody
                TemplateFiller(body, "attachments", mailDetails)
                TemplateFiller(body, "HostDetails", GetSystemInformation())
                TemplateFiller(body, "NetworkDetails", GetNetworkIP())
                outGoingMail.Body = body
                email.Close(Outlook.OlInspectorClose.olDiscard)
                outGoingMail.Display()
            End If
        Catch ex As Exception
            PrepareErrorMail(ex)
        End Try
    End Sub
#End Region

#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Function GetImage(ByVal control As IRibbonControl) As Bitmap
        Return My.Resources.logo
    End Function

    Public Function GetButtonLabel(ByVal control As IRibbonControl) As String
        Return My.Settings.BTN_LABEL
    End Function

    Public Function GetSupertipLabel(ByVal control As IRibbonControl) As String
        Return My.Settings.SUPERTIP_LABEL
    End Function

    Public Function GetGroupLabel(ByVal control As IRibbonControl) As String
        Return My.Settings.GROUP_LABEL
    End Function

    Public Sub BTNclick(ByVal control As IRibbonControl)
        ' Decide via the ID of the control what to do and redirect to the correct method '
        Select Case control.Id
            Case "BTNMailTab"
                ProcessMainWindowMail()
            Case "BTNSendReceiveTab"
                ProcessMainWindowMail()
            Case "BTNNewMailTab"
                ProcessWindowMail()
            Case "BTNReadTab"
                ProcessWindowMail()
            Case Else
                ' Well tbh this cannot be the case '
                Throw New NotSupportedException("The given control ID " + control.Id + " is not defined")
        End Select
    End Sub


#End Region

#Region "Helpers"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), System.StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As System.IO.StreamReader = New System.IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function
    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        ' Function to return the correct IDs for the different tabs '
        Select Case ribbonID
            Case "Microsoft.Outlook.Explorer"
                Return GetResourceText("GOVCERT_Outlook_Addins.RibbonHome.xml")
            Case "Microsoft.Outlook.Mail.Read"
                Return GetResourceText("GOVCERT_Outlook_Addins.RibbonReadMail.xml")
            Case "Microsoft.Outlook.Mail.Compose"
                Return GetResourceText("GOVCERT_Outlook_Addins.RibbonNewMail.xml")
            Case Else
                Return Nothing
        End Select
    End Function
#End Region

End Class
