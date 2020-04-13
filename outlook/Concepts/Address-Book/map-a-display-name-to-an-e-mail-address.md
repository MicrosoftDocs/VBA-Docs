---
title: Map a Display Name to an Email Address
ms.prod: outlook
ms.assetid: ac4e12f8-ea0f-02df-5ce9-23a1c7eda8e0
ms.date: 06/08/2019
localization_priority: Normal
---


# Map a Display Name to an Email Address

This topic shows a code sample in Visual Basic for Applications (VBA) that takes a display name and tries to map it to an email address known to the messaging system in the current session. 

For each Outlook session, the transport provider defines a set of address book containers that the messaging system can deliver messages to. Each address book container corresponds to an address list in Outlook. If a display name is defined in the set of address book containers, the display name can be resolved in the current session and there exists an entry in an address list that maps to this display name. Note that entries in an address list can be of various types, including an Exchange user and Exchange distribution list.

In this code sample, the function `ResolveDisplayNameToSMTP` uses the display name "Dan Wilson" as an example. It first tries to verify that the display name is defined in an address list by creating a **[Recipient](../../../api/Outlook.Recipient.md)** object based on this display name and then calling ** [Recipient.Resolve](../../../api/Outlook.Recipient.Resolve.md)**. If the name is resolved, then `ResolveDisplayNameToSMTP` uses the **[AddressEntry](../../../api/Outlook.AddressEntry.md)** object that is mapped to the **Recipient** object to further obtain the type and, if possible, the email address:


- If the type of the **AddressEntry** object is an Exchange user, `ResolveDisplayNameToSMTP` calls ** [AddressEntry.GetExchangeUser](../../../api/Outlook.AddressEntry.GetExchangeUser.md)** to obtain the corresponding **[ExchangeUser](../../../api/Outlook.ExchangeUser.md)** object. ** [ExchangeUser.PrimarySmtpAddress](../../../api/Outlook.ExchangeUser.PrimarySmtpAddress.md)** provides the email address that maps to the display name.

- If the **AddressEntry** object is an Exchange distribution list, `ResolveDisplayNameToSMTP` calls ** [AddressEntry.GetExchangeDistributionList](../../../api/Outlook.AddressEntry.GetExchangeDistributionList.md)** to obtain an **[ExchangeDistributionList](../../../api/Outlook.ExchangeDistributionList.md)** object. ** [ExchangeDistributionList.PrimarySmtpAddress](../../../api/Outlook.ExchangeDistributionList.PrimarySmtpAddress.md)** provides the email address that maps to the display name.

```vb
Sub ResolveDisplayNameToSMTP() 
 Dim oRecip As Outlook.Recipient 
 Dim oEU As Outlook.ExchangeUser 
 Dim oEDL As Outlook.ExchangeDistributionList 
 
 Set oRecip = Application.Session.CreateRecipient("Dan Wilson") 
 oRecip.Resolve 
 If oRecip.Resolved Then 
 Select Case oRecip.AddressEntry.AddressEntryUserType 
 Case OlAddressEntryUserType.olExchangeUserAddressEntry 
 Set oEU = oRecip.AddressEntry.GetExchangeUser 
 If Not (oEU Is Nothing) Then 
 Debug.Print oEU.PrimarySmtpAddress 
 End If 
 Case OlAddressEntryUserType.olExchangeDistributionListAddressEntry 
 Set oEDL = oRecip.AddressEntry.GetExchangeDistributionList 
 If Not (oEDL Is Nothing) Then 
 Debug.Print oEDL.PrimarySmtpAddress 
 End If 
 End Select 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]