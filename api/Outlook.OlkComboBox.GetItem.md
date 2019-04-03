---
title: OlkComboBox.GetItem method (Outlook)
keywords: vbaol11.chm1000224
f1_keywords:
- vbaol11.chm1000224
ms.prod: outlook
api_name:
- Outlook.OlkComboBox.GetItem
ms.assetid: 650fa823-fbb9-9013-86af-4f55367475c3
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkComboBox.GetItem method (Outlook)

Obtains a  **String** that represents an item at the specified location in the list of the combo box control.


## Syntax

_expression_. `GetItem` (_Index_)

_expression_ A variable that represents an [OlkComboBox](Outlook.OlkComboBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|

## Return value

A  **String** value that represents the item at the specified location in the list.


## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](Outlook.OlkComboBox.ListCount.md)** -1), then an out-of-bounds error will be returned.


## See also


[OlkComboBox Object](Outlook.OlkComboBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]