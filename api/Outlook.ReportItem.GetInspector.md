---
title: ReportItem.GetInspector property (Outlook)
keywords: vbaol11.chm1649
f1_keywords:
- vbaol11.chm1649
ms.prod: outlook
api_name:
- Outlook.ReportItem.GetInspector
ms.assetid: 2a9ec97b-56c5-f93c-eb42-7ddb93a4697e
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.GetInspector property (Outlook)

Returns an **[Inspector](Outlook.Inspector.md)** object that represents an inspector initialized to contain the specified item. Read-only.


## Syntax

_expression_. `GetInspector`

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Remarks

This property is useful for returning an **Inspector** object in which to display the item, as opposed to using the **[Application.ActiveInspector](Outlook.Application.ActiveInspector.md)** method and setting the **[Inspector.CurrentItem](Outlook.Inspector.CurrentItem.md)** property. If an **Inspector** object already exists for the item, the **GetInspector** property will return that **Inspector** object instead of creating a new one.


## See also


[ReportItem Object](Outlook.ReportItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]