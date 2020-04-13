---
title: TaskRequestUpdateItem.MessageClass property (Outlook)
keywords: vbaol11.chm1936
f1_keywords:
- vbaol11.chm1936
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.MessageClass
ms.assetid: 2e9f8234-115c-bc65-ed12-fd86ac0acfa2
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]