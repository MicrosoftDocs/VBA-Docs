---
title: JournalItem.MessageClass property (Outlook)
keywords: vbaol11.chm1246
f1_keywords:
- vbaol11.chm1246
ms.prod: outlook
api_name:
- Outlook.JournalItem.MessageClass
ms.assetid: 1a47a08f-d7ba-5627-dfae-c918c74074c4
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.MessageClass property (Outlook)

Returns or sets a **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[JournalItem Object](Outlook.JournalItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]