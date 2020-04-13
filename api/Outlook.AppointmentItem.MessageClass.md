---
title: AppointmentItem.MessageClass property (Outlook)
keywords: vbaol11.chm857
f1_keywords:
- vbaol11.chm857
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.MessageClass
ms.assetid: e98318d9-72e9-0914-83c6-3a05f544874f
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]