---
title: MailItem.MessageClass property (Outlook)
keywords: vbaol11.chm1309
f1_keywords:
- vbaol11.chm1309
ms.prod: outlook
api_name:
- Outlook.MailItem.MessageClass
ms.assetid: 93194a21-dbec-ebfa-ae5d-d4f287ebb2bd
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[MailItem Object](Outlook.MailItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]