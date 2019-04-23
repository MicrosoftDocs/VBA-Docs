---
title: ControlFormat.RemoveItem method (Excel)
keywords: vbaxl10.chm630075
f1_keywords:
- vbaxl10.chm630075
ms.prod: excel
api_name:
- Excel.ControlFormat.RemoveItem
ms.assetid: 351c2333-9e8c-90a6-90a9-839f43184bb8
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.RemoveItem method (Excel)

Removes one or more items from a list box or combo box.


## Syntax

_expression_.**RemoveItem** (_Index_, _Count_)

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The number of the first item to be removed. Valid values are from 1 to the number of items in the list (returned by the **ListCount** property).|
| _Count_|Optional| **Variant**|The number of items to be removed, starting at item _Index_. If this argument is omitted, one item is removed. If _Index_ + _Count_ exceeds the number of items in the list, all items from _Index_ through the end of the list are removed without an error.|

## Remarks

If the specified object has a fill range defined for it, this method fails.

Use the **[RemoveAllItems](Excel.ControlFormat.RemoveAllItems.md)** method to remove all entries from a list box or combo box.


## Example

This example removes the selected item from a list box. If `Shapes(2)` doesn't represent a list box, this example fails.

```vb
Set lbcf = Worksheets(1).Shapes(2).ControlFormat 
lbcf.RemoveItem lbcf.ListIndex
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]