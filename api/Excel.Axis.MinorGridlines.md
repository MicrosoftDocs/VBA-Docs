---
title: Axis.MinorGridlines property (Excel)
keywords: vbaxl10.chm561092
f1_keywords:
- vbaxl10.chm561092
ms.prod: excel
api_name:
- Excel.Axis.MinorGridlines
ms.assetid: 5725fdb3-05de-e555-5734-cbc64c6a2068
ms.date: 04/13/2019
localization_priority: Normal
---


# Axis.MinorGridlines property (Excel)

Returns a **[Gridlines](Excel.Gridlines(object).md)** object that represents the minor gridlines for the specified axis. Only axes in the primary axis group can have gridlines. Read-only.


## Syntax

_expression_.**MinorGridlines**

_expression_ A variable that represents an **[Axis](Excel.Axis(object).md)** object.


## Example

This example sets the color of the minor gridlines for the value axis on Chart1.

```vb
With Charts("Chart1").Axes(xlValue) 
 If .HasMinorGridlines Then 
 .MinorGridlines.Border.ColorIndex = 5 'set color to blue 
 End If 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]