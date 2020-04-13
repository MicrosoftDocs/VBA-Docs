---
title: DocumentItem.MessageClass property (Outlook)
keywords: vbaol11.chm1198
f1_keywords:
- vbaol11.chm1198
ms.prod: outlook
api_name:
- Outlook.DocumentItem.MessageClass
ms.assetid: 635ba15e-cacc-4e3e-0824-8ca4dfca2a82
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[DocumentItem Object](Outlook.DocumentItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]