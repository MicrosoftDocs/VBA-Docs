---
title: Shape.ControlFormat property (Excel)
keywords: vbaxl10.chm636128
f1_keywords:
- vbaxl10.chm636128
ms.prod: excel
api_name:
- Excel.Shape.ControlFormat
ms.assetid: e874098f-ea8c-93ff-f746-a0d568bec5b5
ms.date: 05/14/2019
localization_priority: Normal
---


# Shape.ControlFormat property (Excel)

Returns a **[ControlFormat](Excel.ControlFormat.md)** object that contains Microsoft Excel control properties. Read-only.


## Syntax

_expression_.**ControlFormat**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


## Example

This example removes the selected item from a list box. If `Shapes(2)` doesn't represent a list box, this example fails.

```vb
Set lbcf = Worksheets(1).Shapes(2).ControlFormat 
lbcf.RemoveItem lbcf.ListIndex
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]