---
title: TaskRequestDeclineItem.MessageClass property (Outlook)
keywords: vbaol11.chm1838
f1_keywords:
- vbaol11.chm1838
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.MessageClass
ms.assetid: 8d244971-e28f-fa88-a115-fad220f3f400
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]