---
title: ControlFormat.Min property (Excel)
keywords: vbaxl10.chm630086
f1_keywords:
- vbaxl10.chm630086
ms.prod: excel
api_name:
- Excel.ControlFormat.Min
ms.assetid: e5b70b54-5304-d013-2398-128609ddb7af
ms.date: 04/23/2019
localization_priority: Normal
---


# ControlFormat.Min property (Excel)

Returns or sets the minimum value of a scroll bar or spinner range. The scroll bar or spinner won't take on values less than this minimum value. Read/write **Long**.


## Syntax

_expression_.**Min**

_expression_ An expression that returns a **[ControlFormat](Excel.ControlFormat.md)** object.


## Return value

Long


## Remarks

The value of the **Min** property must be less than the value of the **[Max](Excel.ControlFormat.Max.md)** property.


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