---
title: OlkListBox.SetItem method (Outlook)
keywords: vbaol11.chm1000269
f1_keywords:
- vbaol11.chm1000269
ms.prod: outlook
api_name:
- Outlook.OlkListBox.SetItem
ms.assetid: 95232643-c547-f553-1d92-0f3fead18de9
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkListBox.SetItem method (Outlook)

Sets the item at the specified location in the list to the specified value.


## Syntax

_expression_. `SetItem`( `_Index_` , `_Item_` )

_expression_ A variable that represents an [OlkListBox](Outlook.OlkListBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|A zero-based value that specifies the location of an item in the list.|
| _Item_|Required| **String**|The value to be used to update the list at the specified location.|

## Remarks

If  _Index_ is outside the range of the allowed values (between zero and **[ListCount](Outlook.OlkListBox.ListCount.md)** -1), then an out-of-bounds error will be returned.


## See also


[OlkListBox Object](Outlook.OlkListBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]