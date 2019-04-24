---
title: ControlFormat.ListIndex property (Excel)
keywords: vbaxl10.chm630083
f1_keywords:
- vbaxl10.chm630083
ms.prod: excel
api_name:
- Excel.ControlFormat.ListIndex
ms.assetid: 34df9efc-e53b-58fd-31b1-4ae592d3d9a8
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.ListIndex property (Excel)

Returns or sets the index number of the currently selected item in a list box or combo box. Read/write **Long**.


## Syntax

_expression_.**ListIndex**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Remarks

You cannot use this property with multiselect list boxes.


## Example

This example removes the selected item from a list box. If `Shapes(2)` doesn't represent a list box, this example fails.

```vb
Set lbcf = Worksheets(1).Shapes(2).ControlFormat 
lbcf.RemoveItem lbcf.ListIndex
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]