---
title: NoteItem.GetInspector property (Outlook)
keywords: vbaol11.chm1482
f1_keywords:
- vbaol11.chm1482
ms.prod: outlook
api_name:
- Outlook.NoteItem.GetInspector
ms.assetid: 80e5bdc5-8161-afa7-6aab-65356fc5d2ea
ms.date: 06/08/2017
localization_priority: Normal
---


# NoteItem.GetInspector property (Outlook)

Returns an  **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [NoteItem](Outlook.NoteItem.md) object.


## Remarks

This property is useful for returning an  **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[NoteItem Object](Outlook.NoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]