---
title: OlkComboBox.AddItem method (Outlook)
keywords: vbaol11.chm1000230
f1_keywords:
- vbaol11.chm1000230
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.AddItem
ms.assetid: 8670b0ba-b715-e00d-0eb9-fa7279ae52b7
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.AddItem method (Outlook)

Adds an item to the list, optionally specifying an index for the new item to appear in the list.


## Syntax

_expression_.**AddItem** (_ItemText_, _Index_)

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ItemText_|Required| **String**|Value to be added to the list in the combo box.|
| _Index_|Optional| **Long**|A 0-based value that specifies the order of the new item in the list.|

## Remarks

If the value of  _Index_ is equal to or larger than the number of elements in the list, the new item will be added to the end of the list.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]