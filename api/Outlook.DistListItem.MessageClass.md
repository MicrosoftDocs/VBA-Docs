---
title: DistListItem.MessageClass property (Outlook)
keywords: vbaol11.chm1129
f1_keywords:
- vbaol11.chm1129
ms.prod: outlook
api_name:
- Outlook.DistListItem.MessageClass
ms.assetid: a719fb30-fee2-24c1-77aa-4650b90bf426
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[DistListItem Object](Outlook.DistListItem.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]