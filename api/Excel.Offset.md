---
title: Offset property (Excel Graph)
keywords: vbagr10.chm5252675
f1_keywords:
- vbagr10.chm5252675
ms.prod: excel
api_name:
- Excel.Offset
ms.assetid: f2f00d51-2a85-aa9c-4361-69f4534cd8e5
ms.date: 04/11/2019
localization_priority: Normal
---


# Offset property (Excel Graph)

Returns or sets the distance between each of the levels of labels, and the distance between the first level and the axis line. The default is 100, which represents the spacing between the axis labels and axis line. The value can be an integer from 0 to 1000, relative to the size of the font of the axis label. Read/write **Long**.

## Syntax

_expression_.**Offset**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example doubles the existing tick-mark spacing on the value axis in _myChart_ if the offset is less than 500.

```vb
With myChart.Axes(xlCategory).TickLabels 
 If .Offset < 500 then 
 .Offset = .Offset * 2 
 End If 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]