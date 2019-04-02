---
title: AutoFormatRules.Add method (Outlook)
keywords: vbaol11.chm2719
f1_keywords:
- vbaol11.chm2719
ms.prod: outlook
api_name:
- Outlook.AutoFormatRules.Add
ms.assetid: 23edea51-416a-22f3-f62e-61f69de5a753
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRules.Add method (Outlook)

Creates a new  **[AutoFormatRule](Outlook.AutoFormatRule.md)** object and appends it to the **[AutoFormatRules](Outlook.AutoFormatRules.md)** collection.


## Syntax

_expression_.**Add** (_Name_)

_expression_ A variable that represents an [AutoFormatRules](Outlook.AutoFormatRules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new formatting rule.|

## Return value

An  **AutoFormatRule** object that represents the new formatting rule.


## Remarks

Duplicate names for  **AutoFormatRule** objects are allowed in the **AutoFormatRules** collection. A maximum of 25 custom formatting rules can be added to the collection. Built-in formatting rules are not counted against that limit.


## See also


[AutoFormatRules Object](Outlook.AutoFormatRules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]