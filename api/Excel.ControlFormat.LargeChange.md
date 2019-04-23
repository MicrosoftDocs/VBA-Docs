---
title: ControlFormat.LargeChange property (Excel)
keywords: vbaxl10.chm630078
f1_keywords:
- vbaxl10.chm630078
ms.prod: excel
api_name:
- Excel.ControlFormat.LargeChange
ms.assetid: 2e47bd4f-59dc-d620-14f0-e4ecdfb4eb78
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.LargeChange property (Excel)

Returns or sets the amount that the scroll box increments or decrements for a page scroll (when the user clicks in the scroll bar body region). Read/write **Long**.


## Syntax

_expression_.**LargeChange**

_expression_ A variable that represents a **[ControlFormat](Excel.ControlFormat.md)** object.


## Example

This example creates a scroll bar and sets its linked cell, minimum, maximum, large change, and small change values.

```vb
Set sb = Worksheets(1).Shapes.AddFormControl(xlScrollBar, _ 
 Left:=10, Top:=10, Width:=10, Height:=200) 
With sb.ControlFormat 
 .LinkedCell = "D1" 
 .Max = 100 
 .Min = 0 
 .LargeChange = 10 
 .SmallChange = 2 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]