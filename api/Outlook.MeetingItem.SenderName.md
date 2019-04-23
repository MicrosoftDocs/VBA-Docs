---
title: MeetingItem.SenderName property (Outlook)
keywords: vbaol11.chm1450
f1_keywords:
- vbaol11.chm1450
ms.prod: outlook
api_name:
- Outlook.MeetingItem.SenderName
ms.assetid: 07dd4ff2-36cd-cfbd-3b48-08e60f0aed78
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.SenderName property (Outlook)

Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.


## Syntax

_expression_. `SenderName`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagSenderName**.

If you wish to retrieve the fully qualified email address of the sender, use the  **[SenderEmailAddress](Outlook.MeetingItem.SenderEmailAddress.md)** property.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]