---
title: AutoScaling property (Excel Graph)
keywords: vbagr10.chm65643
f1_keywords:
- vbagr10.chm65643
ms.prod: excel
api_name:
- Excel.AutoScaling
ms.assetid: f132291c-e356-eea5-0ef5-0e4def8d4832
ms.date: 04/09/2019
localization_priority: Normal
---


# AutoScaling property (Excel Graph)

**True** if Graph scales a 3D chart so that it's closer in size to the equivalent 2D chart. The **[RightAngleAxes](excel.rightangleaxes.md)** property must be **True**. Read/write **Boolean**.

## Syntax

_expression_.**AutoScaling**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example automatically scales the chart. The example should be run on a 3D chart.

```vb
With myChart 
 .RightAngleAxes = True 
 .AutoScaling = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]