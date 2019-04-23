---
title: RemoteItem.GetInspector property (Outlook)
keywords: vbaol11.chm1597
f1_keywords:
- vbaol11.chm1597
ms.prod: outlook
api_name:
- Outlook.RemoteItem.GetInspector
ms.assetid: 0f8e0621-7094-afd5-8913-9f42d55765e0
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.GetInspector property (Outlook)

Returns an  **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]