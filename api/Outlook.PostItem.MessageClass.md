---
title: PostItem.MessageClass property (Outlook)
keywords: vbaol11.chm1528
f1_keywords:
- vbaol11.chm1528
ms.prod: outlook
api_name:
- Outlook.PostItem.MessageClass
ms.assetid: 4f5064a7-0de0-025b-56f9-3c29c4741e5a
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[PostItem Object](Outlook.PostItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]