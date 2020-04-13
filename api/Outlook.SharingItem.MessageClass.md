---
title: SharingItem.MessageClass property (Outlook)
keywords: vbaol11.chm612
f1_keywords:
- vbaol11.chm612
ms.prod: outlook
api_name:
- Outlook.SharingItem.MessageClass
ms.assetid: d2991917-120f-9d69-156f-793e67f45ed9
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.

The default value for this property is  `IPM.Sharing`.


## See also


[SharingItem Object](Outlook.SharingItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]