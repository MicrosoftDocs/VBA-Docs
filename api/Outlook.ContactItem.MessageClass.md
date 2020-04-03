---
title: ContactItem.MessageClass property (Outlook)
keywords: vbaol11.chm945
f1_keywords:
- vbaol11.chm945
ms.prod: outlook
api_name:
- Outlook.ContactItem.MessageClass
ms.assetid: 3d6594b7-8abe-9e49-64e0-be3062807e34
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MessageClass property (Outlook)

Returns or sets a  **String** representing the message class for the Outlook item. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[ContactItem Object](Outlook.ContactItem.md)




[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]