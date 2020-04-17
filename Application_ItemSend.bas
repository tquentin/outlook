'-------------------------------------------- Code Start below this line --------------------------------------------
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
'
' Modified from
' https://www.slipstick.com/how-to-outlook/prevent-sending-messages-to-wrong-email-address/
'   "Check for different domains"
'
' Modified by Quentin Chung
' Modified for Macroview
' Date: 2019 Jun 05
' version 1.1
' Date: 2020 Apr 09
' version 1.2 Modified to prompt once only - use of dictionary data type from MS Scripting Runtime
' to enable: in Developer windows, Tools->Reference, enable "Microsoft Scripting Runtime"
' Checking of internal emails - email subject with "[INTERNAL]" and expected no external domain recipients
'
    Dim la_OrigRecipientsList As Outlook.Recipients
    Dim lo_OrigRecipient As Outlook.Recipient
    Dim lo_PropertyAccessor As Outlook.PropertyAccessor
    Dim lc_PromptMSG As String
    Dim ln_Answer As String
    Dim lc_RecipientEmailAddress As String
    Dim lc_MyDomain As String
    Dim lc_UserAddress As String
    Dim lc_RecipientDomainName As String    ' used in v1.1 & v1.2
    Dim ln_CountI As Integer    ' used in v1.1 & v1.2
    Dim ld_RecipientDomains As New Scripting.Dictionary     ' new in v1.2
    Dim lc_PromptMSGPre As String
    Dim lc_PromptMSGPost As String
    Dim lc_PromptMSGTitle As String
    Dim lc_InternalString As String

    Const PR_SMTP_ADDRESS As String = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
     
    ' non-exchange
    ' lc_UserAddress = Session.CurrentUser.Address
    ' use for exchange accounts
    lc_UserAddress = Session.CurrentUser.AddressEntry.GetExchangeUser.PrimarySmtpAddress
    lc_MyDomain = Right(lc_UserAddress, Len(lc_UserAddress) - InStrRev(lc_UserAddress, "@"))

    ' lc_MyDomain is the sender's email domain
    ' ld_RecipientDomains is the "dictionary" of recipients' email domain generated from Item.Recipients after below For-loop
    ' Collect Recipient Domains list which <> lc_MyDomain

    Set la_OrigRecipientsList = Item.Recipients
    lc_OrigSubject = Item.Subject

    lc_PromptMSGPost = " Do you still wish to send?"
    lc_PromptMSGTitle = "Multiple Recipient Domain Checking"
    lc_InternalString = "[internal]"

    If InStr(LCase(lc_OrigSubject), LCase(lc_InternalString)) > 0 Then
        lc_PromptMSGPre = "This INTERNAL email is being sent to people at "
    Else
        lc_PromptMSGPre = "This email is being sent to people at "
    End If

    For Each lo_OrigRecipient In la_OrigRecipientsList
        Set lo_PropertyAccessor = lo_OrigRecipient.PropertyAccessor
      
        lc_RecipientEmailAddress = LCase(lo_PropertyAccessor.GetProperty(PR_SMTP_ADDRESS))
        lc_RecipientDomainName = Right(lc_RecipientEmailAddress, Len(lc_RecipientEmailAddress) - InStrRev(lc_RecipientEmailAddress, "@"))

        If InStr(LCase(lc_RecipientDomainName), LCase(lc_MyDomain)) = 0 And Not ld_RecipientDomains.Exists(LCase(lc_RecipientDomainName)) Then
            ld_RecipientDomains.Add LCase(lc_RecipientDomainName), LCase(lc_RecipientDomainName)
        End If

    Next

    If (ld_RecipientDomains.count >= 2 And InStr(LCase(lc_OrigSubject), LCase(lc_InternalString)) = 0) Or (ld_RecipientDomains.count > 0 And InStr(LCase(lc_OrigSubject), LCase(lc_InternalString)) > 0) Then
        ' Prompt as (RecipientDomains have > 1 + NOT INTERNAL)==>multiple recipients in >1 domains OR (RecipientDomains have > 0 + INTERNAL)==>INTERNAL email but has external recipient(s)
        lc_PromptMSG = ld_RecipientDomains.Items(0)
        For ln_CountI = 1 To ld_RecipientDomains.count - 1
            lc_PromptMSG = lc_PromptMSG & vbNewLine & ld_RecipientDomains.Items(ln_CountI)
        Next
        lc_PromptMSG = lc_PromptMSGPre & vbNewLine & vbNewLine & lc_PromptMSG & vbNewLine & vbNewLine & lc_PromptMSGPost
        ln_Answer = MsgBox(lc_PromptMSG, vbYesNoCancel + vbExclamation + vbMsgBoxSetForeground, lc_PromptMSGTitle)
        Select Case ln_Answer
            Case vbCancel
                Cancel = True
                Exit Sub ' stops checking for matches
            Case vbNo
                Cancel = True
        End Select
    End If
End Sub
'--------------------------------------------  Code End above this line  --------------------------------------------
