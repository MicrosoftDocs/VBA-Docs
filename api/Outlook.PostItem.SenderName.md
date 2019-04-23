---
title: PostItem.SenderName property (Outlook)
keywords: vbaol11.chm1550
f1_keywords:
- vbaol11.chm1550
ms.prod: outlook
api_name:
- Outlook.PostItem.SenderName
ms.assetid: cee9b0ac-1528-1387-48db-b31d58d691ca
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.SenderName property (Outlook)

Returns a  **String** indicating the display name of the sender for the Outlook item. Read-only.


## Syntax

_expression_. `SenderName`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagSenderName**.

If you wish to retrieve the fully qualified email address of the sender, use the  **[SenderEmailAddress](Outlook.PostItem.SenderEmailAddress.md)** property.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]