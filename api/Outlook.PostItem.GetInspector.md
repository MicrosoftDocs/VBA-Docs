---
title: PostItem.GetInspector property (Outlook)
keywords: vbaol11.chm1524
f1_keywords:
- vbaol11.chm1524
ms.prod: outlook
api_name:
- Outlook.PostItem.GetInspector
ms.assetid: 705fe03b-2ff4-8ed8-e3c2-fb7d52444169
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.GetInspector property (Outlook)

Returns an  **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]