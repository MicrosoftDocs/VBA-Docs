---
title: AppointmentItem.SendUsingAccount property (Outlook)
keywords: vbaol11.chm923
f1_keywords:
- vbaol11.chm923
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.SendUsingAccount
ms.assetid: c3a73b32-c2e1-cb32-35e3-e278f78700ad
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.SendUsingAccount property (Outlook)

Returns or sets an **[Account](Outlook.Account.md)** object that represents the account under which the **[AppointmentItem](Outlook.AppointmentItem.md)** is to be sent. Read/write.


## Syntax

_expression_. `SendUsingAccount`

 _expression_ An expression that returns a [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

The **SendUsingAccount** property can be used to specify the account that should be used to send the **AppointmentItem** when the **[Send](Outlook.TaskItem.Send(method).md)** method is called. This property returns **Null** (**Nothing** in Visual Basic) if the account specified for the **AppointmentItem** no longer exists.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]