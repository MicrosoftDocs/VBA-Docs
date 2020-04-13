---
title: ContactItem.GetInspector property (Outlook)
keywords: vbaol11.chm941
f1_keywords:
- vbaol11.chm941
ms.prod: outlook
api_name:
- Outlook.ContactItem.GetInspector
ms.assetid: d1f8530f-f797-413f-92cb-d0e8215de0e4
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]