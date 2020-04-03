---
title: Rules.Item method (Outlook)
keywords: vbaol11.chm2159
f1_keywords:
- vbaol11.chm2159
ms.prod: outlook
api_name:
- Outlook.Rules.Item
ms.assetid: fe696181-9f61-0eb7-9634-5f7c007f1606
ms.date: 06/08/2017
localization_priority: Normal
---


# Rules.Item method (Outlook)

Obtains a **[Rule](Outlook.Rule.md)** object specified by _Index_ , which is either a numerical index into the **[Rules](Outlook.Rules.md)** collection or the rule name.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [Rules](Outlook.Rules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a 1-based  **long** value representing an index into the **Rules** collection, or a **string** name representing the value of the default property of a rule, **[Rule.Name](Outlook.Rule.Name.md)**.|

## Return value

A **Rule** object that matches the rule specified by _Index_.


## Remarks

Returns an error when the rule cannot be found in the collection.


## See also


[Rules Object](Outlook.Rules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]