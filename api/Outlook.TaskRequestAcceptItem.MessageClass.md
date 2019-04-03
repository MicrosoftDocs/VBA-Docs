---
title: TaskRequestAcceptItem.MessageClass property (Outlook)
keywords: vbaol11.chm1789
f1_keywords:
- vbaol11.chm1789
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.MessageClass
ms.assetid: 817ffe01-109d-5121-96c9-d4738b1dfd91
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.MessageClass property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]