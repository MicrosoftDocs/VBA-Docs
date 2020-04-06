---
title: RemoteItem.MessageClass property (Outlook)
keywords: vbaol11.chm1601
f1_keywords:
- vbaol11.chm1601
ms.prod: outlook
api_name:
- Outlook.RemoteItem.MessageClass
ms.assetid: cdb17ebc-ea8a-31b1-ef32-e9e4dda872c7
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.MessageClass property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]