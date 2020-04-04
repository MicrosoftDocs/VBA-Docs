---
title: SharingItem.GetInspector property (Outlook)
keywords: vbaol11.chm608
f1_keywords:
- vbaol11.chm608
ms.prod: outlook
api_name:
- Outlook.SharingItem.GetInspector
ms.assetid: 960f9b66-35dc-54ab-13c3-9ea54802bccf
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified **[SharingItem](Outlook.SharingItem.md)**. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]