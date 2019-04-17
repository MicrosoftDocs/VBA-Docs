---
title: Rules.Remove method (Outlook)
keywords: vbaol11.chm2162
f1_keywords:
- vbaol11.chm2162
ms.prod: outlook
api_name:
- Outlook.Rules.Remove
ms.assetid: 6d4bb971-b38a-0434-1b6a-8892689549d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Rules.Remove method (Outlook)

Removes from the  **Rules** collection a **Rule** object specified by _Index_ , which is either a numerical index into the **[Rules](Outlook.Rules.md)** collection or the rule name.


## Syntax

_expression_.**Remove** (_Index_)

_expression_ A variable that represents a [Rules](Outlook.Rules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a  **long** value representing an index into the **Rules** collection, or a **string** name representing the value of the default property of a rule, **[Rule.Name](Outlook.Rule.Name.md)**.|

## Remarks

Returns an error when the rule cannot be found in the collection.


## See also


[Rules Object](Outlook.Rules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]