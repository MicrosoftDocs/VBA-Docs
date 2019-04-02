---
title: FormDescription.MessageClass property (Outlook)
keywords: vbaol11.chm191
f1_keywords:
- vbaol11.chm191
ms.prod: outlook
api_name:
- Outlook.FormDescription.MessageClass
ms.assetid: 51ab2c14-de92-b029-e5b8-2e158a626319
ms.date: 06/08/2017
localization_priority: Normal
---


# FormDescription.MessageClass property (Outlook)

Returns a  **String** representing the message class for the **[FormDescription](Outlook.FormDescription.md)** object. Read-only.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents a [FormDescription](Outlook.FormDescription.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[FormDescription Object](Outlook.FormDescription.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]