---
title: RuleConditions.Item method (Outlook)
keywords: vbaol11.chm2301
f1_keywords:
- vbaol11.chm2301
ms.prod: outlook
api_name:
- Outlook.RuleConditions.Item
ms.assetid: 2fc986a5-e77a-e8c9-b8bf-4af85720a771
ms.date: 06/08/2017
localization_priority: Normal
---


# RuleConditions.Item method (Outlook)

Obtains a **[RuleCondition](Outlook.RuleCondition.md)** object specified by _Index_ which is a numerical index into the **[RuleConditions](Outlook.RuleConditions.md)** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [RuleConditions](Outlook.RuleConditions.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A 1-based numerical value that reflects the ordinal position of a rule condition within the  **RuleConditions** collection. For example, the index value of the first rule condition in the collection is 1, and the index value of the second rule condition is 2.|

## Return value

A **RuleCondition** object that represents the specified object.


## Remarks

The **RuleConditions** collection object is a fixed collection. It contains **RuleCondition** objects or objects derived from **RuleCondition**. You cannot add or remove items from this collection, but you can index into the collection to enumerate the rule condition items, and set the **[Enabled](Outlook.RuleCondition.Enabled.md)** property of the rule condition. When using **Item** to enumerate the collection, you can enumerate _Index_ from 1 to **[Count](Outlook.RuleConditions.Count.md)**.


## See also


[RuleConditions Object](Outlook.RuleConditions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]