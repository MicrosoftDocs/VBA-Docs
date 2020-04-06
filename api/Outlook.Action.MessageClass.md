---
title: Action.MessageClass property (Outlook)
keywords: vbaol11.chm16
f1_keywords:
- vbaol11.chm16
ms.prod: outlook
api_name:
- Outlook.Action.MessageClass
ms.assetid: a1a1eaeb-2772-babc-18ba-28ce9a66500b
ms.date: 06/08/2017
localization_priority: Normal
---


# Action.MessageClass property (Outlook)

Returns or sets a  **String** representing the message class for the **[Action](Outlook.Action.md)**. Read/write.


## Syntax

_expression_. `MessageClass`

_expression_ A variable that represents an [Action](Outlook.Action.md) object.


## Remarks

This property corresponds to the MAPI property  **PidTagMessageClass**. The **MessageClass** property links the item to the form on which it is based. When an item is selected, Outlook uses the message class to locate the form and expose its properties, such as **Reply** commands.


## See also


[Action Object](Outlook.Action.md)



[Item Types and Message Classes](../outlook/Concepts/Forms/item-types-and-message-classes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]