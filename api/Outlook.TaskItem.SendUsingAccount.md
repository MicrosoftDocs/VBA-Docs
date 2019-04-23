---
title: TaskItem.SendUsingAccount property (Outlook)
keywords: vbaol11.chm1768
f1_keywords:
- vbaol11.chm1768
ms.prod: outlook
api_name:
- Outlook.TaskItem.SendUsingAccount
ms.assetid: 711382c3-1003-cf0e-2f29-fc3f9d4320a8
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.SendUsingAccount property (Outlook)

Returns or sets an  **[Account](Outlook.Account.md)** object that represents the account under which the **[TaskItem](Outlook.TaskItem.md)** object is to be sent. Read/write.


## Syntax

_expression_. `SendUsingAccount`

 _expression_ An expression that returns a [TaskItem](Outlook.TaskItem.md) object.


## Remarks

The  **SendUsingAccount** property can be used to specify the account that should be used to send the **TaskItem** object when the **[Send](Outlook.TaskItem.Send(method).md)** method is called. This property returns **Null** (**Nothing** in Visual Basic) if the account specified for the **TaskItem** object no longer exists.


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]