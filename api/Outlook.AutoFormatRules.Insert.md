---
title: AutoFormatRules.Insert method (Outlook)
keywords: vbaol11.chm2720
f1_keywords:
- vbaol11.chm2720
ms.prod: outlook
api_name:
- Outlook.AutoFormatRules.Insert
ms.assetid: fb2f4c41-b4f7-fa70-3f44-ee6b818a46ee
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRules.Insert method (Outlook)

Creates a new  **[AutoFormatRule](Outlook.AutoFormatRule.md)** object and inserts it at the specified index within the **[AutoFormatRules](Outlook.AutoFormatRules.md)** collection.


## Syntax

_expression_.**Insert** (_Name_, _Index_)

_expression_ A variable that represents an [AutoFormatRules](Outlook.AutoFormatRules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new object.|
| _Index_|Required| **Variant**|Either the index number at which to insert the new object, or a value used to match the  **[Name](Outlook.AutoFormatRule.Name.md)** property value of an object in the collection at where the new object is to be inserted.|

## Return value

An  **AutoFormatRule** object that represents the new formatting rule.


## Remarks

This method cannot be used to insert custom formatting rules between or ahead of built-in formatting rules.

Duplicate names for  **AutoFormatRule** objects are allowed in the **AutoFormatRules** collection. A maximum of 25 custom formatting rules can be added to the collection. Built-in formatting rules are not counted against that limit.


## See also


[AutoFormatRules Object](Outlook.AutoFormatRules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]