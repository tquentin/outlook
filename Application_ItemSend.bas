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
' version 1.0
' Date: 2020 Apr 09
' version 1.1 Modified to prompt once only - use of dictionary data type from MS Scripting Runtime
' to enable: in Developer windows, Tools->Reference, enable "Microsoft Scripting Runtime"
'
    Dim la_OrigRecipientsList As Outlook.Recipients
    Dim lo_OrigRecipient As Outlook.Recipient
    Dim lo_PropertyAccessor As Outlook.PropertyAccessor
    Dim lc_PromptMSG As String
    Dim ln_Answer As String
    Dim lc_RecipientEmailAddress As String
    Dim lc_MyDomain As String
    Dim lc_UserAddress As String
    Dim lc_RecipientDomainName As String    ' used in v1.0 & v1.1
    Dim ln_CountI As Integer    ' used in v1.0 & v1.1
    Dim ld_RecipientDomains As New Scripting.Dictionary     ' new in v1.1

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
    For Each lo_OrigRecipient In la_OrigRecipientsList
        Set lo_PropertyAccessor = lo_OrigRecipient.PropertyAccessor
      
        lc_RecipientEmailAddress = LCase(lo_PropertyAccessor.GetProperty(PR_SMTP_ADDRESS))
        lc_RecipientDomainName = Right(lc_RecipientEmailAddress, Len(lc_RecipientEmailAddress) - InStrRev(lc_RecipientEmailAddress, "@"))

        If InStr(lc_RecipientDomainName, lc_MyDomain) = 0 And Not ld_RecipientDomains.Exists(lc_RecipientDomainName) Then
            ld_RecipientDomains.Add lc_RecipientDomainName, lc_RecipientDomainName
        End If

    Next

	If ld_RecipientDomains.count > 1 Then
		' Prompt as RecipientDomains have > 1 ==> multiple recipients in >1 domains
		lc_PromptMSG = ld_RecipientDomains.Items(0)
		For ln_CountI = 1 To ld_RecipientDomains.count - 1
			lc_PromptMSG = lc_PromptMSG & vbNewLine & ld_RecipientDomains.Items(ln_CountI)
		Next
		lc_PromptMSG = "This email is being sent to people at " & vbNewLine & vbNewLine & lc_PromptMSG & vbNewLine & vbNewLine & " Do you still wish to send?"
		ln_Answer = MsgBox(lc_PromptMSG, vbYesNoCancel + vbExclamation + vbMsgBoxSetForeground, "Multiple Recipient Domain Checking")
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
