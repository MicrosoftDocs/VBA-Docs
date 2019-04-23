---
title: TaskRequestItem.MessageClass property (Outlook)
keywords: vbaol11.chm1887
f1_keywords:
- vbaol11.chm1887
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.MessageClass
ms.assetid: 078d8ef9-ea60-f27c-ad68-da945f5b8fc8
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.MessageClass property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]