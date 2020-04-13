---
title: MeetingItem.MessageClass property (Outlook)
keywords: vbaol11.chm1417
f1_keywords:
- vbaol11.chm1417
ms.prod: outlook
api_name:
- Outlook.MeetingItem.MessageClass
ms.assetid: 0e7f893f-4de3-06c6-32e0-c815f9af35d5
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]