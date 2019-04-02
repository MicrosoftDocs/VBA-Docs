---
title: Store.GetRules method (Outlook)
keywords: vbaol11.chm810
f1_keywords:
- vbaol11.chm810
ms.prod: outlook
api_name:
- Outlook.Store.GetRules
ms.assetid: 06048799-e162-68f9-17c2-d80c25e2c55e
ms.date: 06/08/2017
localization_priority: Normal
---


# Store.GetRules method (Outlook)

Returns a  **[Rules](Outlook.Rules.md)** collection object that contains the **[Rule](Outlook.Rule.md)** objects defined for the current session.


## Syntax

_expression_. `GetRules`

_expression_ A variable that represents a [Store](Outlook.Store.md) object.


## Return value

A  **Rules** collection object that represents the set of **Rules** defined for the current session.


## Remarks

Calling  **GetRules** can be an expensive operation in terms of performance on slow connections to an Exchange server.

The order of the  **Rule** objects in the collection returned from **GetRules** follows that of **[Rule.ExecutionOrder](Outlook.Rule.ExecutionOrder.md)** with **ExecutionOrder** equal 1 being the first **Rule** in the collection and **ExecutionOrder** equal **[Rules.Count](Outlook.Rules.Count.md)** being the last **Rule** in the collection.


## See also


[Store Object](Outlook.Store.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]