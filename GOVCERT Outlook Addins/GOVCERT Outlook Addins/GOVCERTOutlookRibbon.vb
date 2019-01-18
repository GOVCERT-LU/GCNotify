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

Imports Microsoft.Office.Core

<System.Runtime.InteropServices.ComVisible(True)>
Public Class GOVCERTOutlookRibbon
    Implements IRibbonExtensibility

    Private ribbon As IRibbonUI
    Private PS_PUBLIC_STRINGS As String = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}"

    Public Sub New()
        ' Constructor '

    End Sub

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

    Private Function GetSubject(mail As Microsoft.Office.Interop.Outlook.MailItem) As String
        ' Removes unicode emojis from the subject -> this is due to an RT / MySQL bug '
        Dim unicodeString As String = mail.Subject
        Dim unicodeEncoding As System.Text.Encoding = System.Text.Encoding.Unicode
        Dim asciiEncoding As System.Text.Encoding = System.Text.Encoding.ASCII

        ' Convert the string into a byte array. '
        Dim unicodeAsBytes As Byte() = unicodeEncoding.GetBytes(unicodeString)
        ' Perform the conversion from one encoding to the other. 
        Dim asciiBytes As Byte() = System.Text.Encoding.Convert(unicodeEncoding, asciiEncoding, unicodeAsBytes)

        ' Convert the new byte array into a char array and then into a string. 
        Dim asciiChars(asciiEncoding.GetCharCount(asciiBytes, 0, asciiBytes.Length) - 1) As Char
        asciiEncoding.GetChars(asciiBytes, 0, asciiBytes.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)
        Return asciiString
    End Function

    Private Function CreateNewMail() As Microsoft.Office.Interop.Outlook.MailItem
        ' Creates the new mail to send out '
        Dim email As Microsoft.Office.Interop.Outlook.MailItem
        email = CType(Globals.ThisAddIn.Application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem), Microsoft.Office.Interop.Outlook.MailItem)
        PopulateNewMail(email)
        Return email
    End Function

    Private Sub PopulateNewMail(email As Microsoft.Office.Interop.Outlook.MailItem)
        ' Sets the values for the new mail '
        email.Subject = My.Resources.Tag + "-" + My.Resources.SuspectTag
        email.Body = My.Resources.NewMailBody
        email.To = My.Resources.These
        email.PropertyAccessor.SetProperty(PS_PUBLIC_STRINGS + "/X-GC-Notify-Version", My.Resources.Version)
    End Sub

    Private Function GenerateAttachment(outgoingMail As Microsoft.Office.Interop.Outlook.MailItem, attachmentMail As Microsoft.Office.Interop.Outlook.MailItem, counter As Integer) As String
        ' Attaches the selected mails to the New Mail '
        Dim attachMail As Boolean = True
        If attachmentMail.Subject.Contains("SPAM") Then
            'This means that the mail was already tagged as spam then show the dialog for confirmation'
            Dim message As String = My.Resources.SPAMDialogText
            message = message.Replace("{{emailAddress}}", attachmentMail.SenderEmailAddress)
            message = message.Replace("{{emailSubject}}", GetSubject(attachmentMail))
            Dim result As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(message, "WARNING - Email already Tagged", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning)
            If result = System.Windows.Forms.DialogResult.No Then
                attachMail = False
            Else
                ' Set forced header '
                ' Set X-Headers '
                outgoingMail.PropertyAccessor.SetProperty(PS_PUBLIC_STRINGS + "/X-GC-Notify-Force-Send", "true")
            End If
        End If
        If attachMail Then
            outgoingMail.Attachments.Add(attachmentMail, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue)
            Return "[" + GetSubject(attachmentMail) + "] from: " + attachmentMail.SenderEmailAddress
        Else
            Return Nothing
        End If
    End Function

    Private Function NewMail(mail As Microsoft.Office.Interop.Outlook.MailItem) As Boolean
        ' Takes care of double sendings '
        If String.IsNullOrEmpty(mail.Subject) Then
            PopulateNewMail(mail)
            Return False
        Else
            If mail.Subject.StartsWith(My.Resources.Tag) Then
                ' Show dialog an do nothing '
                System.Windows.Forms.MessageBox.Show(My.Resources.NewResendError, "ERROR - GOVCERT.LU Addins", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Return Nothing
            Else
                If String.IsNullOrEmpty(mail.Body) Then
                    PopulateNewMail(mail)
                    Return False
                Else
                    Dim result As System.Windows.Forms.DialogResult = System.Windows.Forms.MessageBox.Show(My.Resources.OverWriteConfirm, "Do you want to proceed?", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question)
                    If result = System.Windows.Forms.DialogResult.Yes Then
                        PopulateNewMail(mail)
                    End If
                End If
            End If
        End If
        Return Nothing
    End Function

    Public Sub SendMail(control As IRibbonControl)
        Dim subjectTag As String = My.Resources.Tag + My.Resources.SuspectTag
        Dim subject As String = Nothing
        Dim body As String = My.Resources.SuspectBody
        Dim attachment As String = Nothing
        Dim attachmentText As String = Nothing
        Dim counter As Integer = 0

        Dim createMail As Boolean = True
        Dim outGoingMail As Microsoft.Office.Interop.Outlook.MailItem = Nothing

        ' check where the button was called from '
        'Yes this could be done other wise '

        Dim inspector As Microsoft.Office.Interop.Outlook.Inspector = Globals.ThisAddIn.Application.ActiveInspector()
        If TypeOf inspector Is Microsoft.Office.Interop.Outlook.Inspector Then
            If TypeOf inspector.CurrentItem Is Microsoft.Office.Interop.Outlook.MailItem Then
                Dim mail As Microsoft.Office.Interop.Outlook.MailItem = CType(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
                ' First check if it Is Not a New mail '
                ' if the EntryID Is empty the mail was Not saved, hence Is a New one '
                If String.IsNullOrEmpty(mail.EntryID) Then
                    createMail = NewMail(mail)
                Else
                    ' Take case of double sendings '
                    Try
                        mail.PropertyAccessor.GetProperty(PS_PUBLIC_STRINGS + "/X-GC-Notify-Version")
                        ' Show dialog an do nothing '
                        System.Windows.Forms.MessageBox.Show(My.Resources.ResendError, "ERROR - GOVCERT.LU Addins", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                        Return
                    Catch ex As System.Runtime.InteropServices.COMException
                        If outGoingMail Is Nothing Then
                            outGoingMail = CreateNewMail()
                        End If
                        attachment = GenerateAttachment(outGoingMail, mail, counter)
                        If Not attachment Is Nothing Then
                            subject = attachment
                            attachmentText = attachmentText + System.Environment.NewLine + attachment
                            counter = 1
                        End If
                    End Try
                End If

            End If
        Else
            Try
                Dim selectedMails As Microsoft.Office.Interop.Outlook.Selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection
                If outGoingMail Is Nothing Then
                    outGoingMail = CreateNewMail()
                End If
                For Each email As Microsoft.Office.Interop.Outlook.MailItem In selectedMails
                    attachment = GenerateAttachment(outGoingMail, email, counter)
                    If Not attachment Is Nothing Then
                        subject = attachment
                        attachmentText = attachmentText + System.Environment.NewLine + attachment
                        counter = counter + 1
                    End If
                Next

            Catch ex As System.Runtime.InteropServices.COMException
                System.Windows.Forms.MessageBox.Show(My.Resources.NoSelectionError, "ERROR - GOVCERT.LU Addins", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                Return

            End Try
        End If

        If createMail Then
            If outGoingMail Is Nothing Then
                outGoingMail = CreateNewMail()
            End If
            outGoingMail.Body = body.Replace("{{attachments}}", attachmentText)
            If counter > 0 Then
                If counter = 1 Then
                    outGoingMail.Subject = subjectTag + " - " + subject
                Else
                    outGoingMail.Subject = subjectTag + " - Multiple Emails"
                End If
                outGoingMail.Display()
            Else
                System.Windows.Forms.MessageBox.Show("No Email(s) selected", "ERROR - GOVCERT.LU Addins", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation)
                Return
            End If
        End If
        Return

    End Sub


#Region "Ribbon Callbacks"
    'Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub

    Public Function GetImage(ByVal imageName As String) As stdole.IPictureDisp
        Select Case imageName
            Case "gcnotify"
                Return PictureConverter.IconToPictureDisp(My.Resources.gcnotify)
            Case Else
                Return Nothing
        End Select

    End Function

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

#End Region

End Class
