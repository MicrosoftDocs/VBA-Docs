---
title: MeetingItem.SendUsingAccount property (Outlook)
keywords: vbaol11.chm3509
f1_keywords:
- vbaol11.chm3509
ms.prod: outlook
api_name:
- Outlook.MeetingItem.SendUsingAccount
ms.assetid: 81713c7b-dfb0-eb91-b017-82b427bee823
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.SendUsingAccount property (Outlook)

Returns or sets an  **[Account](Outlook.Account.md)** object that represents the account to use to send the **[MeetingItem](Outlook.MeetingItem.md)**. Read/write.


## Syntax

_expression_. `SendUsingAccount`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

You can use the  **SendUsingAccount** property to specify the account that the **Send** method uses to send the **MeetingItem**. This property returns **Null** (**Nothing** in Visual Basic) if the account specified for the **MeetingItem** no longer exists.

This property is read-only if the  **MeetingItem** is a received item, or if the **MeetingItem** has already been sent (that is, the **[Sent](Outlook.MeetingItem.Sent.md)** property of the object is set to **True**).


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]