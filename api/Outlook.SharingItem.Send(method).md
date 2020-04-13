---
title: SharingItem.Send method (Outlook)
keywords: vbaol11.chm672
f1_keywords:
- vbaol11.chm672
ms.prod: outlook
api_name:
- Outlook.SharingItem.Send
ms.assetid: 54f92175-0e99-f96a-56de-5fc66d97d80f
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Send method (Outlook)

Sends the  **[SharingItem](Outlook.SharingItem.md)**.


## Syntax

_expression_. `Send`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The **Send** method sends an item using the default account specified for the session. In a session where multiple Microsoft Exchange accounts are defined in the profile, the first Exchange account added to the profile is the primary Exchange account, and is also the default account for the session. To specify a different account to send an item, set the **[SendUsingAccount](Outlook.SharingItem.SendUsingAccount.md)** property to the desired **[Account](Outlook.Account.md)** object and then call the **Send** method.

Certain sharing providers may have restrictions on the type of recipients allowed. When this method is called, some providers will attempt to set access control list (ACL) entries on the folder for each recipient included in the  **SharingItem**. If an error occurs while attempting to set ACLs for any recipient, this method raises an error and the **SharingItem** is not sent to any of the recipients.

An error occurs if the  **[BCC](Outlook.SharingItem.BCC.md)** or **[CC](Outlook.SharingItem.CC.md)** properties are set for a **SharingItem** using an Exchange sharing context.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]