---
title: HasLegend property (Excel Graph)
keywords: vbagr10.chm65589
f1_keywords:
- vbagr10.chm65589
ms.prod: excel
api_name:
- Excel.HasLegend
ms.assetid: b4dbef39-9d83-2f6e-fe06-8ca38cceeeec
ms.date: 04/11/2019
localization_priority: Normal
---


# HasLegend property (Excel Graph)

**True** if the chart has a legend. Read/write **Boolean**.

## Syntax

_expression_.**HasLegend**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example turns on the legend for the chart, and then sets the legend font color to blue.

```vb
With myChart 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]