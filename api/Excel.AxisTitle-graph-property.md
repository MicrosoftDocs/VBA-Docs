---
title: AxisTitle property (Excel Graph)
keywords: vbagr10.chm65618
f1_keywords:
- vbagr10.chm65618
ms.prod: excel
api_name:
- Excel.AxisTitle
ms.assetid: 2fa829a9-e414-6826-32c5-27189b913409
ms.date: 06/08/2017
localization_priority: Normal
---


# AxisTitle property (Excel Graph)

Returns an AxisTitle object that represents the title of the specified axis. Read-only AxisTitle object.

## Syntax

_expression_. `AxisTitle`

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example adds an axis label to the category axis in  `myChart`.


```vb
With myChart.Axes(xlCategory) 
    .HasTitle = True 
    .AxisTitle.Text = "July Sales" 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]