---
title: JournalItem.GetInspector property (Outlook)
keywords: vbaol11.chm1242
f1_keywords:
- vbaol11.chm1242
ms.prod: outlook
api_name:
- Outlook.JournalItem.GetInspector
ms.assetid: 49d173ba-e4fd-e9c4-12b4-423a4c60ec46
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]