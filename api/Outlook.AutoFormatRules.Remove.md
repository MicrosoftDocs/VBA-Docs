---
title: AutoFormatRules.Remove method (Outlook)
keywords: vbaol11.chm2721
f1_keywords:
- vbaol11.chm2721
ms.prod: outlook
api_name:
- Outlook.AutoFormatRules.Remove
ms.assetid: 91db9890-c8cd-81bd-fd10-4137072ca2b5
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoFormatRules.Remove method (Outlook)

Removes an object from the collection.


## Syntax

_expression_.**Remove** (_Index_)

_expression_ A variable that represents an [AutoFormatRules](Outlook.AutoFormatRules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either the index number of the object, or a value used to match the  **[Name](Outlook.AutoFormatRule.Name.md)** property value of an object in the collection.|

## Remarks

If the name of a formatting rule to be removed is specified in  _Index_, this method removes the first  **[AutoFormatRule](Outlook.AutoFormatRule.md)** object that matches the specified name.

Built-in formatting rules cannot be removed from the collection.


## See also


[AutoFormatRules Object](Outlook.AutoFormatRules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]