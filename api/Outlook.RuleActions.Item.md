---
title: RuleActions.Item method (Outlook)
keywords: vbaol11.chm2183
f1_keywords:
- vbaol11.chm2183
ms.prod: outlook
api_name:
- Outlook.RuleActions.Item
ms.assetid: d37a3f0c-0273-e4c2-21e5-661484244671
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleActions.Item method (Outlook)

Obtains a **[RuleAction](Outlook.RuleAction.md)** object specified by _Index_ which is a numerical index into the **[RuleActions](Outlook.RuleActions.md)** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [RuleActions](Outlook.RuleActions.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A 1-based numerical value that reflects the ordinal position of a rule action within the  **RuleActions** collection. For example, the index value of the first rule action in the collection is 1, and the index value of the second rule action is 2.|

## Return value

A **RuleAction** object that matches the rule action specified by _Index_.


## Remarks

The **RuleActions** collection object is a fixed collection. It contains **RuleAction** objects or objects derived from **RuleAction**. You cannot add or remove items from this collection, but you can use **Item** to enumerate the rule action items, and set the **[Enabled](Outlook.RuleAction.Enabled.md)** property of the rule action. When using **Item** to enumerate the collection, you can enumerate _Index_ from 1 to **[Count](Outlook.RuleActions.Count.md)**.


## See also


[RuleActions Object](Outlook.RuleActions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]