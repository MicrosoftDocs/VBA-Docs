---
title: TaskItem.MessageClass property (Outlook)
keywords: vbaol11.chm1701
f1_keywords:
- vbaol11.chm1701
ms.prod: outlook
api_name:
- Outlook.TaskItem.MessageClass
ms.assetid: e5deb86e-ad13-32f0-8dd8-802e7cc539aa
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[TaskItem Object](Outlook.TaskItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]